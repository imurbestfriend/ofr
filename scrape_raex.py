#!/usr/bin/env python3
"""Collect attributed company, business and industry information from RA Expert."""
import argparse
from dataclasses import asdict, dataclass
from datetime import datetime
import json
import os
from pathlib import Path
import re
import sys
from urllib.parse import urljoin, urlsplit, urlunsplit

from bs4 import BeautifulSoup
from dotenv import load_dotenv
from llm_config import build_messages, get_base_url, get_completion_options, get_model
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://raexpert.ru"
DATE_RE = re.compile(r"\b(\d{2}\.\d{2}\.\d{4})\b")
BOND_RE = re.compile(r"облигац|долгов\w* инструмент", re.I)
SUMMARY_PROMPT = """Составь подробное исследование компании ТОЛЬКО по переданным источникам.
Сохрани максимум полезных уникальных подробностей, не добавляй внешних знаний и догадок.
У каждого фактического утверждения укажи источник [S1], [S2] и т.д. Используй только
переданные идентификаторы. Различай юридическое лицо, группу, дочерние и совместные
предприятия. У чисел сохраняй значение, единицу, период и периметр. Подозрительные
единицы не исправляй молча. Исторические сведения помечай датами, противоречия указывай.
Новые сведения имеют приоритет. Сопоставь единицы измерения разных публикаций и
явно отметь несогласованность. Для финансовых потоков сохраняй LTM/год/полугодие.
Отделяй факты от оценок и прогнозов агентства.
Отраслевые тенденции не приписывай компании. Излагай своими словами, без длинных цитат.
Разделы: компания и география; бизнес-модель, сырьё, продукты и производственная цепочка;
активы, предприятия и мощности; рынки, клиенты, поставщики и конкуренция; инвестиции и
стратегия; отрасль (спрос, цены, регулирование, барьеры, тенденции); финансовые и
операционные показатели по периодам; преимущества, зависимости и риски; пробелы и
противоречия в источниках. Отсутствующие сведения не выдумывай. Это подробное досье,
а не краткое резюме: раскрывай конкретные названия, связи, проекты и показатели.
"""


def clean(text):
    return " ".join(text.split())


def site_url(href, base=BASE_URL):
    """Canonicalize mirrors and fragments; keep navigation on RA Expert."""
    parsed = urlsplit(urljoin(base, href))
    if parsed.scheme not in ("http", "https") or not parsed.hostname:
        return None
    if parsed.hostname != "raexpert.ru" and not parsed.hostname.endswith(".raexpert.ru"):
        return None
    path = parsed.path or "/"
    if path.startswith(("/database/companies/", "/researches/", "/topics/")):
        path = path.rstrip("/") + "/"
    return urlunsplit(("https", "raexpert.ru", path, "", ""))


def parse_date(text):
    match = DATE_RE.search(text)
    if match:
        try:
            return datetime.strptime(match[1], "%d.%m.%Y").date().isoformat()
        except ValueError:
            pass
    return ""


@dataclass
class Source:
    id: str
    url: str
    title: str
    published: str
    kind: str
    text: str

    def render(self):
        return (f"[{self.id}] {self.title}\nURL: {self.url}\n"
                f"Дата публикации: {self.published or 'не установлена'}\n"
                f"Тип: {self.kind}\n\n{self.text}")


class Site:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers["User-Agent"] = "CompanyResearch/1.0 (public company reports)"
        self.session.mount("https://", HTTPAdapter(max_retries=Retry(
            total=2, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])))
        self.cache = {}

    def get(self, url):
        url = site_url(url)
        if not url:
            raise ValueError("Ожидалась ссылка на raexpert.ru")
        if url not in self.cache:
            response = self.session.get(url, timeout=(10, 45))
            response.raise_for_status()
            if not site_url(response.url):
                raise ValueError("Перенаправление за пределы raexpert.ru")
            self.cache[url] = response.text
        return BeautifulSoup(self.cache[url], "lxml")

    def search(self, query):
        # Search is a POST form with a session-bound CSRF token, not a GET query.
        soup = self.get(BASE_URL)
        form = soup.select_one('form:has(input[name="search"])')
        data = {i["name"]: i.get("value", "") for i in form.select('input[name]')} if form else {}
        data["search"] = query
        response = self.session.post(BASE_URL + "/search/", data=data,
                                     headers={"Referer": BASE_URL + "/"}, timeout=(10, 45))
        response.raise_for_status()
        return BeautifulSoup(response.text, "lxml")


def page_title(soup):
    heading = soup.find("h1") or soup.find("title")
    if not heading:
        return ""
    heading = BeautifulSoup(str(heading), "lxml")
    for item in heading.select(".b-more, .b-pdf-button-header"):
        item.decompose()
    return clean(heading.get_text(" ")).removesuffix(" | Эксперт РА")


def company_details(soup):
    details = {}
    for row in soup.select(".b-table__item"):
        label, value = row.select_one(".b-table__title"), row.select_one(".b-table__data")
        if label and value:
            details[clean(label.get_text())] = clean(value.get_text(" "))
    return details


def name_words(name):
    return [w for w in re.findall(r"[а-яёa-z0-9]+", name.lower())
            if w not in {"пао", "ао", "оао", "зао", "ооо", "публичное", "акционерное", "общество"}]


def search_company(site, name, inn=None):
    soup = site.search(inn or name)
    candidates = {}
    for a in soup.select('#companies a[href*="/database/companies/"]'):
        url = site_url(a["href"])
        if url:
            candidates[url] = clean(a.get_text(" "))
    if inn:
        candidates = {u: n for u, n in candidates.items()
                      if company_details(site.get(u)).get("ИНН") == inn}
    else:
        exact = {u: n for u, n in candidates.items() if name_words(n) == name_words(name)}
        if exact:
            candidates = exact
    if len(candidates) != 1:
        choices = "\n".join(f"  {n}: {u}" for u, n in candidates.items())
        raise ValueError("Компания не найдена или результат неоднозначен. "
                         "Укажите --inn либо --company-url.\n" + choices)
    return next(iter(candidates))


def linked_pages(soup, prefix):
    links = {}
    for a in soup.select(f'a[href*="{prefix}"]'):
        url = site_url(a["href"])
        if not url or not urlsplit(url).path.startswith(prefix):
            continue
        parent = a.find_parent("tr") or a.parent
        links[url] = (clean(a.get_text(" ")), parse_date(parent.get_text(" ")))
    return links


def release_candidates(soup):
    candidates = {}
    for block in soup.select(".b-actions__rates"):
        subtitle = block.select_one(".b-actions__subtitle")
        label = clean(subtitle.get_text(" ")) if subtitle else ""
        if BOND_RE.search(label):
            continue
        kind = "esg" if "esg" in label.lower() else "credit"
        for url, (_, published) in linked_pages(block, "/releases/").items():
            candidates[url] = (kind, published)
    for block in soup.select(".b-aside-content__item"):
        for url, (title, _) in linked_pages(block, "/releases/").items():
            if not BOND_RE.search(title):
                candidates.setdefault(url, ("esg" if "esg" in title.lower() else "credit",
                                             parse_date(block.get_text(" "))))
    return sorted(candidates.items(), key=lambda item: item[1][1], reverse=True)


def article_text(soup):
    article = soup.select_one(".b-article__body")
    if not article:
        raise ValueError("Не найден текст статьи (.b-article__body)")
    article = BeautifulSoup(str(article), "lxml")
    for item in article.select("script, style, nav, aside, .-info-contacts, .release-graph-wrapper.-type-range"):
        item.decompose()
    # Remove only the regulatory section in its own column, preserving the
    # financial tables and related reports in the other column.
    for heading in list(article.find_all(["h2", "h3"])):
        if re.search(r"Регуляторное|Контакты|Отказ от ответственности", heading.get_text(), re.I):
            for sibling in list(heading.next_siblings):
                if getattr(sibling, "name", None) in ("h2", "h3"):
                    break
                if getattr(sibling, "get_text", None) and re.match(
                        r"Краткая информация|Ключевые финансовые|Связанные отч[её]ты",
                        clean(sibling.get_text(" ")), re.I):
                    break
                sibling.extract()
            heading.decompose()
    for table in list(article.find_all("table")):
        rows = []
        for row in table.find_all("tr"):
            cells = [clean(c.get_text(" ")).replace("|", "\\|")
                     for c in row.find_all(["td", "th"], recursive=False)]
            if cells:
                rows.append("| " + " | ".join(cells) + " |")
        table.replace_with("\n" + "\n".join(rows) + "\n")
    return re.sub(r"\n{3,}", "\n\n", article.get_text("\n", strip=True)).strip()


def extract_source(soup, url, kind, published=""):
    stamp = soup.select_one(".b-subheader__date")
    return Source("", url, page_title(soup),
                  (parse_date(stamp.get_text()) if stamp else "") or published,
                  kind, article_text(soup))


def industry_stems(industry):
    stems = [w[:5] for w in name_words(industry)
             if len(w) > 3 and w not in {"промышленность", "прочие", "компании", "услуги"}]
    if any(w.startswith("нефт") for w in stems):
        stems.extend(["нефт", "газов", "спг"])
    if any(w.startswith("хим") for w in stems):
        stems.extend(["нефтехим", "полимер"])
    return stems


def collect(site, company_url, credit_limit=6, esg_limit=3, research_limit=4,
            extra_research=(), inn=None):
    card = site.get(company_url)
    details = company_details(card)
    if inn and details.get("ИНН") != inn:
        raise ValueError("ИНН выбранной карточки не совпадает с --inn")
    title = page_title(card)
    if not title or not details.get("ИНН"):
        raise ValueError("Не удалось подтвердить карточку компании и её ИНН")
    sources = [Source("S1", company_url, title, "", "company",
                      "\n".join(f"{k}: {v}" for k, v in details.items()))]
    warnings, research, topics = [], {}, set()
    counts, limits = {"credit": 0, "esg": 0}, {"credit": credit_limit, "esg": esg_limit}
    stems = industry_stems(details.get("Отрасль", ""))

    def discover(soup):
        article = soup.select_one(".b-article__body")
        for url, meta in linked_pages(article or soup, "/researches/").items():
            if any(stem in meta[0].lower() for stem in stems):
                research.setdefault(url, meta)

    for link in card.select('.b-aside-content__source a[href*="/topics/"]'):
        if any(stem in link.get_text().lower() for stem in stems):
            topics.add(site_url(link["href"]))
    for url, (kind, published) in release_candidates(card):
        if counts[kind] >= limits[kind]:
            continue
        try:
            soup = site.get(url)
            source = extract_source(soup, url, kind, published)
            if BOND_RE.search(source.title):
                continue
            if not all(word in name_words(source.title) for word in name_words(title)):
                continue
            discover(soup)
            if len(source.text) < 800 or not re.search(
                    r"Обоснование|бизнес|отрасл|производ|деятельност", source.text, re.I):
                warnings.append(f"Пропущен краткий релиз: {url}")
                continue
            sources.append(source)
            counts[kind] += 1
            print(f"  {kind}: {source.published} {source.title}", file=sys.stderr)
        except (requests.RequestException, ValueError) as exc:
            warnings.append(f"{url}: {exc}")
    if research_limit:
        for url in sorted(topics):
            try:
                discover(site.get(url))
            except (requests.RequestException, ValueError) as exc:
                warnings.append(f"{url}: {exc}")
        industry = details.get("Отрасль", "")
        if industry:
            try:
                discover(site.search(industry))
            except (requests.RequestException, ValueError) as exc:
                warnings.append(f"Поиск исследований по отрасли «{industry}»: {exc}")
    explicit = {site_url(url) for url in extra_research}
    if None in explicit:
        raise ValueError("--research-url должен указывать на raexpert.ru")
    for url in sorted(explicit):
        if not urlsplit(url).path.startswith("/researches/"):
            raise ValueError("--research-url должен указывать на /researches/...")
        research.setdefault(url, ("", ""))
    candidates = sorted(research, key=lambda u: (u in explicit, research[u][1],
                        max(re.findall(r"20\d{2}", u), default=""), u), reverse=True)
    collected = []
    for url in candidates:
        if len(collected) >= research_limit and url not in explicit:
            continue
        # Archived handbooks are not a current industry outlook. Explicit URLs
        # can still be used for historical research.
        years = re.findall(r"20\d{2}", url)
        if url not in explicit and years and int(max(years)) < datetime.now().year - 3:
            continue
        try:
            source = extract_source(site.get(url), url, "industry")
            if (url not in explicit and source.published
                    and int(source.published[:4]) < datetime.now().year - 3):
                continue
            if len(source.text) < 800:
                warnings.append(f"Не найден полный текст исследования: {url}")
                continue
            collected.append(source)
        except (requests.RequestException, ValueError) as exc:
            warnings.append(f"{url}: {exc}")
    sources.extend(sorted(collected, key=lambda s: s.published, reverse=True))
    if not counts["credit"]:
        warnings.append("Полные кредитные релизы компании не найдены.")
    if not collected:
        warnings.append("Релевантные отраслевые исследования не найдены; можно добавить --research-url.")
    for index, source in enumerate(sources, 1):
        source.id = f"S{index}"
    return {"company": title, "company_url": company_url, "details": details,
            "collected_at": datetime.now().astimezone().isoformat(),
            "warnings": warnings, "sources": [asdict(s) for s in sources]}


def summarize(sources, *, company_prompt=None, industry_prompt=None):
    from openai import OpenAI
    client = OpenAI(
        api_key=os.environ["OCTAVA_API_KEY"],
        base_url=get_base_url(),
    )

    def ask(system, text):
        response = client.chat.completions.create(
            model=get_model(), temperature=0, **get_completion_options(),
            messages=build_messages(system, text))
        choice = response.choices[0]
        if choice.finish_reason != "stop" or not choice.message.content:
            raise ValueError(f"Неполный ответ модели: {choice.finish_reason}")
        result = choice.message.content.strip()
        references = set(re.findall(r"\[(S\d+)\]", result))
        if not references or references - set(re.findall(r"\[(S\d+)\]", text)):
            raise ValueError("В ответе модели отсутствуют ссылки или указаны неизвестные источники")
        return result

    def synthesize(group, prompt):
        corpus = "\n\n---\n\n".join(s.render() for s in group)
        if len(corpus) <= 100_000:
            return ask(prompt, corpus)
        batches, current = [], ""
        # Each fragment retains its source metadata; batch small sources together.
        for source in group:
            for offset in range(0, len(source.text), 24_000):
                fragment = Source(source.id, source.url, source.title, source.published,
                                  source.kind, source.text[offset:offset + 24_000]).render()
                if len(current) + len(fragment) > 60_000:
                    batches.append(current)
                    current = ""
                current += "\n\n" + fragment
        if current:
            batches.append(current)
        notes = []
        for index, batch in enumerate(batches, 1):
            print(f"  Извлекаем факты: пакет {index}/{len(batches)}", file=sys.stderr)
            notes.append(ask(SUMMARY_PROMPT + "\nЭто фрагменты источников: извлеки подробные "
                             "структурированные факты для последующего объединения, "
                             "без общего заключения. Сохрани все уникальные подробности.", batch))
        while sum(map(len, notes)) > 100_000:
            batches, current = [], ""
            for note in notes:
                if len(current) + len(note) > 80_000:
                    batches.append(current)
                    current = ""
                current += "\n\n" + note
            if current:
                batches.append(current)
            notes = [ask(prompt + "\nОбъедини заметки, убери повторы, сохрани ссылки.", b)
                     for b in batches]
        return ask(prompt, "\n\n".join(notes))

    # A separate pass prevents sector studies being lost among company releases.
    company = [s for s in sources if s.kind != "industry"]
    industry = [s for s in sources if s.kind == "industry"]
    sections = []
    if company:
        sections.append(synthesize(company, company_prompt or SUMMARY_PROMPT +
                                  "\nПри достаточном материале подготовь 1200–2000 слов. "
                                  "Раскрой каждый раздел подробно."))
    if industry:
        sections.append("## Дополнительный отраслевой анализ\n\n" + synthesize(
            industry, industry_prompt or SUMMARY_PROMPT + "\nСоставь ТОЛЬКО отраслевой раздел: спрос и предложение, "
            "продукты, цены, регулирование, инвестиции, риски, сценарии и динамика по годам. "
            "Не описывай компанию без прямого подтверждения. При достаточном материале "
            "подготовь 600–1200 слов. Отметь географию и даты каждого обзора."))
    return "\n\n".join(sections)


def save_report(bundle, path, summary=""):
    warnings = "\n".join(bundle["warnings"]) or "Нет."
    text = (f"КОМПАНИЯ: {bundle['company']}\nСобрано: {bundle['collected_at']}\n"
            f"Карточка: {bundle['company_url']}\n\n=== ОГРАНИЧЕНИЯ СБОРА ===\n{warnings}\n\n")
    if summary:
        text += f"=== ОПИСАНИЕ ===\n{summary}\n\n"
    text += "=== RAEX: ИСТОЧНИКИ ===\n\n" + "\n\n---\n\n".join(Source(**s).render() for s in bundle["sources"])
    path.write_text(text, encoding="utf-8")


def nonnegative(value):
    number = int(value)
    if number < 0:
        raise argparse.ArgumentTypeError("Ожидалось неотрицательное число")
    return number


def main():
    load_dotenv()
    parser = argparse.ArgumentParser(description="Подробное досье компании из материалов raexpert.ru")
    parser.add_argument("name", help="Название компании")
    parser.add_argument("--company-url", help="Прямая ссылка на карточку вместо поиска")
    parser.add_argument("--inn", help="ИНН для проверки компании")
    parser.add_argument("--credit-limit", type=nonnegative, default=6)
    parser.add_argument("--esg-limit", type=nonnegative, default=3)
    parser.add_argument("--research-limit", type=nonnegative, default=4)
    parser.add_argument("--research-url", action="append", default=[], help="Дополнительный обзор (можно повторять)")
    parser.add_argument("--no-summary", action="store_true", help="Только сбор, без OpenAI и API-ключа")
    parser.add_argument("--from-json", type=Path, help="Описание из сохранённого JSON без загрузки сайта")
    parser.add_argument("-o", "--output", type=Path, help="Выходной TXT; рядом сохраняется .sources.json")
    args = parser.parse_args()
    path = args.output or Path(re.sub(r'[/\\]', '_', args.name) + ".txt")
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        if args.from_json:
            bundle = json.loads(args.from_json.read_text(encoding="utf-8"))
        else:
            site = Site()
            url = site_url(args.company_url) if args.company_url else search_company(site, args.name, args.inn)
            if not url or not re.fullmatch(r"/database/companies/[^/]+/?", urlsplit(url).path):
                raise ValueError("--company-url должен указывать на /database/companies/...")
            print(f"Собираем: {url}", file=sys.stderr)
            bundle = collect(site, url, args.credit_limit, args.esg_limit,
                             args.research_limit, args.research_url, args.inn)
        json_path = path.with_suffix(".sources.json")
        json_path.write_text(json.dumps(bundle, ensure_ascii=False, indent=2), encoding="utf-8")
        save_report(bundle, path)  # Preserve the corpus even if GPT fails.
        for warning in bundle["warnings"]:
            print(f"Предупреждение: {warning}", file=sys.stderr)
        print(f"Сохранено источников: {len(bundle['sources'])}; {path}; {json_path}", file=sys.stderr)
        if not args.no_summary:
            print("Генерируем подробное описание...", file=sys.stderr)
            summary = summarize([Source(**s) for s in bundle["sources"]])
            save_report(bundle, path, summary)
            print(f"Описание сохранено: {path}", file=sys.stderr)
    except Exception as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
