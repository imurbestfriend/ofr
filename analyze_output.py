#!/usr/bin/env python3
import argparse
import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Side, Border
from openpyxl.utils import get_column_letter

load_dotenv()

client = OpenAI()

CONCLUSION_PROMPT = """\
Ты — финансовый аналитик. На основе приведённого списка рисков напиши краткий итоговый вывод \
(2–3 предложения) в следующем формате: сначала сильные стороны (факторы, снижающие риск), \
затем слабые стороны (ключевые риски). Пиши одним связным абзацем, без заголовков и маркеров.\
"""

SEVERITY_COLORS = {
    "высокий": "FF4C4C",
    "умеренный": "FFC000",
    "низкий": "70AD47",
}

CHUNK_SIZE = 80_000  # chars per API call

CF_METRICS_NAMES = [
    "Выручка LTM",
    "EBITDA LTM",
    "Рентабельность по EBITDA LTM",
    "Чистая прибыль LTM",
    "Долг/EBITDA",
    "ICR LTM",
    "ISCR LTM",
    "DSCR с долгом LTM",
    "Операционный поток (CFO)",
    "Инвестиционный поток (CFI)",
    "Кап. затраты (CAPEX)",
    "Финансовый поток (CFF)",
    "Погашение долга",
    "Привлечение долга",
    "Чистый денежный поток (NCF)",
    "Денежные средства",
    "Долг",
    "Долг перед Сбером",
]

CF_CONCLUSION_PROMPT = """\
Ты — финансовый аналитик. На основе приведённых финансовых метрик напиши краткий вывод \
(2–3 предложения) о финансовом состоянии компании: ключевые сильные стороны и основные риски. \
Пиши одним связным абзацем, без заголовков и маркеров.\
"""

SYSTEM_PROMPT = """\
Ты — экстрактор данных. Твоя единственная задача — найти и дословно извлечь информацию из \
предоставленного тебе текста. Ты работаешь ТОЛЬКО с тем текстом, который получил. \
Твои собственные знания, обучающие данные и любая внешняя информация полностью запрещены.

Верни ТОЛЬКО валидный JSON-объект со следующими ключами:

{
  "metrics": [
    {"name": "название метрики (Revenue/EBITDA/EBIT/Net Income/Total Debt/Net Debt/Capex/FCF и т.п.)",
     "period": "период точно как в тексте (например, 1П2026, H1 2025, FY2025)",
     "value": "числовое значение скопировано дословно из текста",
     "unit": "единица измерения точно как в тексте",
     "quote": "дословная цитата из текста, из которой взята эта метрика (1-2 предложения)"}
  ],
  "cf_metrics": [
    {"metric": "одно из 18 фиксированных названий (см. ниже)",
     "period": "период точно как в тексте (например, 6м2026, 12м2024, 1П2025, FY2024, 2025)",
     "value": "значение дословно из текста (число или процент)"}
  ],
  "risks": [
    {"name": "краткое название риска (например: Финансовый риск, Риск концентрации, Валютный риск)",
     "category": "категория риска (рыночный/кредитный/операционный/геополитический/репутационный/правовой/другой)",
     "severity": "высокий/умеренный/низкий — твоя оценка серьёзности на основе контекста",
     "risk_metrics": "конкретные метрики или индикаторы риска дословно из текста; если явно не указаны — сформулируй на основе описания",
     "detailed_description": "детальное описание риска — дословно из текста или близко к тексту",
     "mitigation_measures": "меры минимизации риска — дословно из текста; если не указаны — предложи исходя из контекста",
     "source": "extracted — если все ключевые поля (описание, метрики, меры) взяты дословно из текста; generated — если хотя бы одно поле было сформулировано тобой при отсутствии в тексте"}
  ],
  "guidance": [
    {"topic": "тема прогноза",
     "detail": "дословная или близкая к дословной цитата из текста с цифрами и сроками"}
  ],
  "highlights": [
    {"event": "краткое название события",
     "detail": "дословная или близкая к дословной цитата из текста"}
  ]
}

СТРОГИЕ ПРАВИЛА (применяются ко всем полям):
- Используй ТОЛЬКО информацию из предоставленного текста. Никаких данных из обучения.
- Если цифра не указана в тексте явно — не добавляй её. Пиши только то, что написано.
- Значения (value) и цитаты (quote) копируй дословно, не округляй и не интерпретируй.
- Если в фрагменте нет данных для категории — верни пустой список [].
- Ответ — только JSON, без markdown-обёртки, без пояснений вне JSON.

ПРАВИЛА ДЛЯ cf_metrics (только для этого поля):
- Все 18 допустимых метрик — исключительно ФИНАНСОВЫЕ показатели (из P&L, Cash Flow Statement, баланса). НЕ путай с операционными/производственными показателями (добыча, запасы, объём реализации в тоннах/куб.м, численность персонала, дивиденды на акцию).
- Допустимый маппинг: «Выручка LTM» ← выручка от реализации/revenue в руб.; «EBITDA LTM» ← EBITDA; «Рентабельность по EBITDA LTM» ← рентабельность по EBITDA в %; «Чистая прибыль LTM» ← чистая прибыль/прибыль за период в денежных единицах; «Долг/EBITDA» ← коэффициент типа 2.4x; «ICR/ISCR/DSCR с долгом LTM» ← только эти явные коэффициенты; «Операционный поток (CFO)» ← CFO; «Инвестиционный поток (CFI)» ← CFI; «Кап. затраты (CAPEX)» ← капитальные вложения/CAPEX в денежных единицах; «Финансовый поток (CFF)» ← CFF; «Погашение долга», «Привлечение долга» ← явные статьи долга; «Чистый денежный поток (NCF)» ← NCF; «Денежные средства» ← остаток денежных средств; «Долг» ← общий долг/заёмные средства; «Долг перед Сбером» ← только явно.
- Если явного финансового показателя в тексте нет — НЕ включай запись.
"""


def build_chunks(text: str) -> list[str]:
    sections = text.split("\n\n---\n\n")
    chunks, current, current_size = [], [], 0
    for section in sections:
        size = len(section) + (8 if current else 0)
        if current and current_size + size > CHUNK_SIZE:
            chunks.append("\n\n---\n\n".join(current))
            current, current_size = [section], len(section)
        else:
            current.append(section)
            current_size += size
    if current:
        chunks.append("\n\n---\n\n".join(current))
    return chunks


def extract_chunk(chunk: str, index: int, total: int) -> dict:
    print(f"  Chunk {index}/{total} ({len(chunk):,} chars)...", file=sys.stderr)
    response = client.chat.completions.create(
        model="gpt-4o",
        response_format={"type": "json_object"},
        temperature=0,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": chunk},
        ],
        max_tokens=4096,
    )
    raw = response.choices[0].message.content
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        print(f"  Warning: could not parse JSON from chunk {index}", file=sys.stderr)
        return {"metrics": [], "risks": [], "guidance": [], "highlights": []}


def merge_results(results: list[dict]) -> dict:
    merged = {"metrics": [], "cf_metrics": [], "risks": [], "guidance": [], "highlights": []}
    seen_metrics: set[tuple] = set()
    seen_cf: set[tuple] = set()
    seen_risks: set[tuple] = set()

    for r in results:
        for m in r.get("metrics", []):
            key = (m.get("name", "").lower(), m.get("period", "").lower())
            if key not in seen_metrics:
                seen_metrics.add(key)
                merged["metrics"].append(m)

        for item in r.get("cf_metrics", []):
            key = (item.get("metric", ""), item.get("period", "").lower())
            if key not in seen_cf:
                seen_cf.add(key)
                merged["cf_metrics"].append(item)

        for risk in r.get("risks", []):
            key = (risk.get("name", "").lower()[:60], risk.get("category", "").lower())
            if key not in seen_risks:
                seen_risks.add(key)
                merged["risks"].append(risk)

        merged["guidance"].extend(r.get("guidance", []))
        merged["highlights"].extend(r.get("highlights", []))

    return merged


def generate_cf_conclusion(cf_metrics: list[dict]) -> str:
    if not cf_metrics:
        return ""
    summary = "\n".join(
        f"- {item.get('metric', '?')} [{item.get('period', '?')}]: {item.get('value', '')}"
        for item in cf_metrics
    )
    response = client.chat.completions.create(
        model="gpt-4o",
        temperature=0,
        messages=[
            {"role": "system", "content": CF_CONCLUSION_PROMPT},
            {"role": "user", "content": summary},
        ],
        max_tokens=512,
    )
    return response.choices[0].message.content.strip()


def generate_conclusion(risks: list[dict]) -> str:
    if not risks:
        return ""
    risk_summary = "\n".join(
        f"- {r.get('name', '?')} [{r.get('severity', '?')}]: {r.get('detailed_description', '')[:200]}"
        for r in risks
    )
    response = client.chat.completions.create(
        model="gpt-4o",
        temperature=0,
        messages=[
            {"role": "system", "content": CONCLUSION_PROMPT},
            {"role": "user", "content": risk_summary},
        ],
        max_tokens=512,
    )
    return response.choices[0].message.content.strip()



def _sort_period(period_str: str) -> tuple:
    import re
    s = (period_str or "").strip()
    # 6м2026 / 12м2025
    m = re.match(r"^(\d{1,2})м(\d{4})$", s)
    if m:
        return (int(m.group(2)), int(m.group(1)))
    # 1П2025 / 2П2025
    m = re.match(r"^([12])П(\d{4})$", s, re.IGNORECASE)
    if m:
        return (int(m.group(2)), int(m.group(1)) * 6)
    # 1кв2025 / 4кв2025
    m = re.match(r"^(\d)кв(\d{4})$", s, re.IGNORECASE)
    if m:
        return (int(m.group(2)), int(m.group(1)) * 3)
    # H1 2025 / H2 2026
    m = re.match(r"^H([12])\s*(\d{4})$", s, re.IGNORECASE)
    if m:
        return (int(m.group(2)), int(m.group(1)) * 6)
    # FY2024
    m = re.match(r"^FY(\d{4})$", s, re.IGNORECASE)
    if m:
        return (int(m.group(1)), 12)
    # bare year: 2024
    m = re.match(r"^(\d{4})$", s)
    if m:
        return (int(m.group(1)), 12)
    return (9999, 99)


def style_header(ws, row=1):
    fill = PatternFill("solid", fgColor="1F4E79")
    font = Font(bold=True, color="FFFFFF", size=11)
    for cell in ws[row]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 22


def auto_width(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)


def write_excel(data: dict, out_path: str, conclusion: str = "", cf_conclusion: str = "") -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # --- Metrics sheet ---
    ws = wb.create_sheet("Метрики")
    headers = ["Метрика", "Период", "Значение", "Единица", "Цитата из текста"]
    ws.append(headers)
    for m in data["metrics"]:
        ws.append([m.get("name"), m.get("period"), m.get("value"), m.get("unit"), m.get("quote")])
    style_header(ws)
    ws.freeze_panes = "A2"
    auto_width(ws)

    # --- Risks sheet (PDF-style layout) ---
    ws = wb.create_sheet("Риски")

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 45

    thin = Side(style="thin", color="BFBFBF")
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="BDD7EE")
    hdr_font = Font(bold=True, color="1F3864", size=11)

    for col_idx, label in enumerate(["УРОВЕНЬ", "РИСК", "ОПИСАНИЕ", "МЕРЫ МИНИМИЗАЦИИ"], 1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = cell_border
    ws.row_dimensions[1].height = 28
    ws.auto_filter.ref = "A1:D1"
    ws.freeze_panes = "B2"

    _sev_order = {"высокий": 0, "умеренный": 1, "низкий": 2}
    sorted_risks = sorted(
        data["risks"],
        key=lambda r: _sev_order.get((r.get("severity") or "").strip().lower(), 99),
    )
    for r in sorted_risks:
        severity = (r.get("severity") or "").strip().lower()
        row_num = ws.max_row + 1

        ca = ws.cell(row=row_num, column=1, value=(r.get("severity") or "").upper())
        ca.fill = PatternFill("solid", fgColor=SEVERITY_COLORS.get(severity, "BFBFBF"))
        ca.font = Font(bold=True, color="FFFFFF", size=8)
        ca.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90, wrap_text=True)
        ca.border = cell_border

        name = r.get("name") or ""
        category = r.get("category") or ""
        source = r.get("source") or ""
        b_parts = [name]
        if category:
            b_parts.append(category)
        if source:
            b_parts.append(f"[{source}]")
        cb = ws.cell(row=row_num, column=2, value="\n".join(b_parts))
        cb.font = Font(bold=True)
        cb.alignment = Alignment(wrap_text=True, vertical="top")
        cb.border = cell_border

        metrics = r.get("risk_metrics") or ""
        desc = r.get("detailed_description") or ""
        c_parts = []
        if metrics:
            c_parts.append(f"МЕТРИКИ РИСКА:\n{metrics}")
        if desc:
            c_parts.append(f"ДЕТАЛЬНОЕ ОПИСАНИЕ:\n{desc}")
        cc = ws.cell(row=row_num, column=3, value="\n\n".join(c_parts))
        cc.alignment = Alignment(wrap_text=True, vertical="top")
        cc.border = cell_border

        cd = ws.cell(row=row_num, column=4, value=r.get("mitigation_measures") or "")
        cd.alignment = Alignment(wrap_text=True, vertical="top")
        cd.border = cell_border

    if conclusion:
        conclusion_row = ws.max_row + 1
        ws.merge_cells(start_row=conclusion_row, start_column=1, end_row=conclusion_row, end_column=4)
        cell = ws.cell(row=conclusion_row, column=1, value=f"ВЫВОД: {conclusion}")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        cell.border = cell_border
        ws.row_dimensions[conclusion_row].height = 70

    # --- CF-модель sheet ---
    ws = wb.create_sheet("CF-модель")

    thin = Side(style="thin", color="BFBFBF")
    cf_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    navy_fill = PatternFill("solid", fgColor="1F4E79")
    navy_font = Font(bold=True, color="FFFFFF", size=11)
    name_fill = PatternFill("solid", fgColor="F2F2F2")
    row_fills = [PatternFill("solid", fgColor="FFFFFF"), PatternFill("solid", fgColor="F5F5F5")]

    # Build pivot: metric → period → value
    cf_data: dict[str, dict[str, str]] = {m: {} for m in CF_METRICS_NAMES}
    for item in data.get("cf_metrics", []):
        metric = item.get("metric", "")
        period = item.get("period", "")
        value = item.get("value", "")
        if metric in cf_data and period:
            cf_data[metric][period] = value

    # Collect and sort periods
    all_periods = sorted(
        {p for vals in cf_data.values() for p in vals},
        key=_sort_period,
    )

    # Column widths
    ws.column_dimensions["A"].width = 30
    for col_idx in range(2, len(all_periods) + 2):
        ws.column_dimensions[get_column_letter(col_idx)].width = 14

    # Header row
    ws.cell(row=1, column=1, value="Показатели").fill = navy_fill
    ws.cell(row=1, column=1).font = navy_font
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.cell(row=1, column=1).border = cf_border
    for col_idx, period in enumerate(all_periods, 2):
        cell = ws.cell(row=1, column=col_idx, value=period)
        cell.fill = navy_fill
        cell.font = navy_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = cf_border
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "B2"

    # Data rows
    for row_idx, metric in enumerate(CF_METRICS_NAMES, 2):
        alt_fill = row_fills[(row_idx) % 2]
        name_cell = ws.cell(row=row_idx, column=1, value=metric)
        name_cell.font = Font(bold=True, size=10)
        name_cell.fill = name_fill
        name_cell.alignment = Alignment(vertical="center", wrap_text=True)
        name_cell.border = cf_border

        for col_idx, period in enumerate(all_periods, 2):
            val = cf_data[metric].get(period, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill = alt_fill
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = cf_border
        ws.row_dimensions[row_idx].height = 18

    # Вывод row
    if cf_conclusion:
        last_col = max(len(all_periods) + 1, 1)
        vyvod_row = len(CF_METRICS_NAMES) + 2
        ws.merge_cells(start_row=vyvod_row, start_column=1, end_row=vyvod_row, end_column=last_col)
        cell = ws.cell(row=vyvod_row, column=1, value=f"ВЫВОД: {cf_conclusion}")
        cell.fill = navy_fill
        cell.font = navy_font
        cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        cell.border = cf_border
        ws.row_dimensions[vyvod_row].height = 60

    # --- Guidance sheet ---
    ws = wb.create_sheet("Прогнозы Guidance")
    ws.append(["Тема", "Детали"])
    for g in data["guidance"]:
        ws.append([g.get("topic"), g.get("detail")])
    style_header(ws)
    ws.freeze_panes = "A2"
    auto_width(ws)

    # --- Highlights sheet ---
    ws = wb.create_sheet("События Highlights")
    ws.append(["Событие", "Детали"])
    for h in data["highlights"]:
        ws.append([h.get("event"), h.get("detail")])
    style_header(ws)
    ws.freeze_panes = "A2"
    auto_width(ws)

    # Wrap long cells (risks sheet handles alignment per cell)
    for sheet in wb.sheetnames:
        if sheet == "Риски":
            continue
        for row in wb[sheet].iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(out_path)


def main():
    parser = argparse.ArgumentParser(description="Extract key data from output.txt using GPT-4o")
    parser.add_argument("--input", default="output.txt", metavar="FILE", help="Input text file (default: output.txt)")
    parser.add_argument("--out", default="analysis.xlsx", metavar="FILE", help="Output Excel file (default: analysis.xlsx)")
    parser.add_argument("--workers", type=int, default=5, help="Parallel API workers (default: 5)")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: {input_path} not found", file=sys.stderr)
        sys.exit(1)

    text = input_path.read_text(encoding="utf-8")
    print(f"Read {len(text):,} chars from {input_path}", file=sys.stderr)

    chunks = build_chunks(text)
    print(f"Split into {len(chunks)} chunk(s), sending to GPT-4o...", file=sys.stderr)

    results: list[dict] = [None] * len(chunks)
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {
            pool.submit(extract_chunk, chunk, i + 1, len(chunks)): i
            for i, chunk in enumerate(chunks)
        }
        for future in as_completed(futures):
            idx = futures[future]
            results[idx] = future.result()

    data = merge_results(results)
    print(
        f"Extracted: {len(data['metrics'])} metrics, {len(data['cf_metrics'])} cf_metrics, "
        f"{len(data['risks'])} risks, {len(data['guidance'])} guidance, {len(data['highlights'])} highlights",
        file=sys.stderr,
    )

    print("Generating conclusion...", file=sys.stderr)
    conclusion = generate_conclusion(data["risks"])
    cf_conclusion = generate_cf_conclusion(data["cf_metrics"])

    write_excel(data, args.out, conclusion=conclusion, cf_conclusion=cf_conclusion)
    print(f"Saved: {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
