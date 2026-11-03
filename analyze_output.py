#!/usr/bin/env python3
import argparse
import json
import os
import re
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI
from llm_config import build_messages, get_base_url, get_completion_options, get_model
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Side, Border
from openpyxl.utils import get_column_letter

load_dotenv()

client = OpenAI(
    api_key=os.environ["OCTAVA_API_KEY"],
    base_url=get_base_url(),
)

IFRS_PROMPT = """\
Ты — экстрактор финансовой отчётности. Извлеки из текста строки финансовых форм МСФО: \
отчёт о финансовом положении (баланс), отчёт о прибылях и убытках, отчёт о движении \
денежных средств. Верни ТОЛЬКО валидный JSON:

{
  "statements": [
    {
      "type": "название формы (например: Отчёт о финансовом положении)",
      "items": [
        {
          "section": "раздел (например: АКТИВЫ / ОБЯЗАТЕЛЬСТВА / КАПИТАЛ / Выручка и т.п.)",
          "is_total": true/false,
          "name": "название строки дословно из текста",
          "periods": [
            {"period": "период (например: 2025 / 2024 / 1П2025)", "value": "значение дословно из текста или null"}
          ]
        }
      ]
    }
  ]
}

ПРАВИЛА:
- Включай ВСЕ строки форм, включая итоговые (is_total: true для строк "Итого...", "Всего...")
- Значения копируй дословно, включая знак минус если есть
- Если форма в тексте отсутствует — верни пустой список statements: []
- Только JSON, без markdown
"""

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

CHUNK_SIZE = 20_000  # chars per API call
CHUNK_SEPARATOR = "\n\n---\n\n"
MAX_SPLIT_DEPTH = 3
MIN_SPLIT_CHARS = 2_000
MAX_OUTPUT_TOKENS = 32_768
LLM_ERROR_DIR = Path(__file__).resolve().parent / "llm_errors"


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
  "risks": [
    {"name": "краткое название риска (например: Финансовый риск, Риск концентрации, Валютный риск)",
     "category": "категория риска (рыночный/кредитный/операционный/геополитический/репутационный/правовой/другой)",
     "severity": "высокий/умеренный/низкий — твоя оценка серьёзности на основе контекста",
     "risk_metrics": "конкретные метрики или индикаторы риска дословно из текста; если явно не указаны — сформулируй на основе описания",
     "detailed_description": "детальное описание риска — дословно из текста или близко к тексту",
     "mitigation_measures": "меры минимизации риска — дословно из текста; если не указаны — предложи исходя из контекста",
     "source": "extracted — если все ключевые поля (описание, метрики, меры) взяты дословно из текста; generated — если хотя бы одно поле было сформулировано тобой при отсутствии в тексте"}
  ]
}

СТРОГИЕ ПРАВИЛА (применяются ко всем полям):
- Используй ТОЛЬКО информацию из предоставленного текста. Никаких данных из обучения.
- Если цифра не указана в тексте явно — не добавляй её. Пиши только то, что написано.
- Значения (value) и цитаты (quote) копируй дословно, не округляй и не интерпретируй.
- Если в фрагменте нет данных для категории — верни пустой список [].
- Ответ — только JSON, без markdown-обёртки, без пояснений вне JSON.
"""


def build_chunks(text: str) -> list[str]:
    chunks = []
    start = 0
    while start < len(text):
        end = min(start + CHUNK_SIZE, len(text))
        if end < len(text):
            boundary = text.rfind(CHUNK_SEPARATOR, start + CHUNK_SIZE // 2, end)
            if boundary >= 0:
                end = boundary + len(CHUNK_SEPARATOR)
            else:
                boundary = text.rfind("\n", start + CHUNK_SIZE // 2, end)
                if boundary >= 0:
                    end = boundary + 1
        chunks.append(text[start:end])
        start = end
    return chunks


def _parse_extraction(raw: str, kind: str) -> dict:
    if not isinstance(raw, str) or not raw.strip():
        raise ValueError("Пустой ответ модели")
    raw = raw.strip()
    fenced = re.fullmatch(r"```(?:json)?\s*\n?(.*?)\n?```", raw, re.DOTALL | re.IGNORECASE)
    if fenced:
        raw = fenced.group(1).strip()
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("Ожидался JSON-объект")
    fields = ("statements",) if kind == "ifrs" else ("metrics", "risks")
    for field in fields:
        if not isinstance(data.get(field), list) or any(
            not isinstance(item, dict) for item in data[field]
        ):
            raise ValueError(f"Поле {field} должно быть списком объектов")
    if kind == "ifrs":
        for statement in data["statements"]:
            if not isinstance(statement.get("type"), str) or not statement["type"].strip():
                raise ValueError("У формы МСФО отсутствует название type")
            if not isinstance(statement.get("items"), list):
                raise ValueError("Поле items должно быть списком объектов")
            for item in statement["items"]:
                if not isinstance(item, dict) or not isinstance(item.get("name"), str):
                    raise ValueError("Некорректная строка формы МСФО")
                if not isinstance(item.get("periods"), list) or any(
                    not isinstance(period, dict) or "period" not in period or "value" not in period
                    for period in item["periods"]
                ):
                    raise ValueError("Некорректные периоды строки МСФО")
    else:
        for field, required in (("metrics", ("name", "period", "value")), ("risks", ("name",))):
            for item in data[field]:
                if any(not isinstance(item.get(key), str) for key in required):
                    raise ValueError(f"Некорректные обязательные поля {field}")
    return data


def _save_bad_response(response, kind: str, index: int, error: str) -> Path:
    choices = getattr(response, "choices", None) or []
    choice = choices[0] if choices else None
    message = getattr(choice, "message", None)
    usage = getattr(response, "usage", None)
    if usage is not None:
        usage = usage.model_dump() if hasattr(usage, "model_dump") else vars(usage)
    diagnostic = {
        "model": get_model(), "kind": kind, "chunk": index, "error": error,
        "max_tokens": MAX_OUTPUT_TOKENS,
        "reasoning_effort": get_completion_options().get("reasoning_effort"),
        "finish_reason": getattr(choice, "finish_reason", None), "usage": usage,
        "content": getattr(message, "content", None),
        "refusal": getattr(message, "refusal", None),
    }
    LLM_ERROR_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", prefix=f"{kind}-{index}-",
                                     suffix=".json", dir=LLM_ERROR_DIR, delete=False) as output:
        json.dump(diagnostic, output, ensure_ascii=False, indent=2, default=vars)
        return Path(output.name)


def _extract_json(chunk: str, index: int, kind: str, depth: int = 0) -> dict:
    prompt = IFRS_PROMPT if kind == "ifrs" else SYSTEM_PROMPT
    for attempt in range(2):
        response = client.chat.completions.create(
            model=get_model(), response_format={"type": "json_object"},
            temperature=0, **get_completion_options(MAX_OUTPUT_TOKENS),
            messages=build_messages(prompt, chunk),
        )
        choices = getattr(response, "choices", None) or []
        choice = choices[0] if choices else None
        reason = getattr(choice, "finish_reason", None)
        message = getattr(choice, "message", None)
        usage = getattr(response, "usage", None)
        completion_tokens = getattr(usage, "completion_tokens", None)
        budget_exhausted = (isinstance(completion_tokens, int)
                            and completion_tokens >= MAX_OUTPUT_TOKENS)
        try:
            if reason != "stop":
                raise ValueError(f"finish_reason={reason}")
            if getattr(message, "refusal", None):
                raise ValueError("Модель отказалась обрабатывать фрагмент")
            if budget_exhausted:
                raise ValueError(f"Исчерпан бюджет ответа: {completion_tokens}/{MAX_OUTPUT_TOKENS} токенов")
            return _parse_extraction(getattr(message, "content", None), kind)
        except (ValueError, TypeError) as exc:
            diagnostic = _save_bad_response(response, kind, index, str(exc))
            print(f"  {kind} chunk {index}: {exc}; ответ сохранён: {diagnostic}", file=sys.stderr)
            recoverable = reason in ("stop", "length") and not getattr(message, "refusal", None)
            truncated = reason == "length" or budget_exhausted
            if recoverable and not truncated and attempt == 0:
                print(f"  {kind} chunk {index}: повторяем запрос JSON", file=sys.stderr)
                continue
            if recoverable:
                if depth < MAX_SPLIT_DEPTH and len(chunk) >= MIN_SPLIT_CHARS:
                    middle = len(chunk) // 2
                    boundary = chunk.rfind("\n", len(chunk) // 4, middle)
                    if boundary >= 0:
                        middle = boundary + 1
                    print(f"  {kind} chunk {index}: делим фрагмент пополам (уровень {depth + 1})",
                          file=sys.stderr)
                    parts = [_extract_json(part, index, kind, depth + 1)
                             for part in (chunk[:middle], chunk[middle:])]
                    return merge_ifrs_results(parts) if kind == "ifrs" else merge_results(parts)
            raise RuntimeError(
                f"Не удалось получить корректный JSON ({kind}, фрагмент {index}, "
                f"finish_reason={reason}): {exc}. Диагностика: {diagnostic}"
            ) from exc


def extract_ifrs_chunk(chunk: str, index: int, total: int) -> dict:
    print(f"  IFRS chunk {index}/{total} ({len(chunk):,} chars)...", file=sys.stderr)
    return _extract_json(chunk, index, "ifrs")


def merge_ifrs_results(results: list[dict]) -> dict:
    merged_by_type: dict[str, list] = {}
    for r in results:
        for stmt in r.get("statements", []):
            stmt_type = stmt.get("type", "").strip()
            if not stmt_type:
                continue
            if stmt_type not in merged_by_type:
                merged_by_type[stmt_type] = []
            seen = {(i.get("section", ""), i.get("name", "")): i
                    for i in merged_by_type[stmt_type]}
            for item in stmt.get("items", []):
                key = (item.get("section", ""), item.get("name", ""))
                if key not in seen:
                    copied = {**item, "periods": [dict(period) for period in item.get("periods", [])]}
                    merged_by_type[stmt_type].append(copied)
                    seen[key] = copied
                else:
                    periods = seen[key]["periods"]
                    existing = {(p.get("period"), p.get("value")) for p in periods}
                    for period in item.get("periods", []):
                        pair = (period.get("period"), period.get("value"))
                        if pair not in existing:
                            periods.append(dict(period))
                            existing.add(pair)
    return {"statements": [{"type": t, "items": items} for t, items in merged_by_type.items()]}


def extract_chunk(chunk: str, index: int, total: int) -> dict:
    print(f"  Chunk {index}/{total} ({len(chunk):,} chars)...", file=sys.stderr)
    return _extract_json(chunk, index, "metrics")


def merge_results(results: list[dict]) -> dict:
    merged = {"metrics": [], "risks": []}
    seen_metrics: set[tuple] = set()
    seen_risks: set[tuple] = set()

    for r in results:
        for m in r.get("metrics", []):
            key = (m.get("name", "").lower(), m.get("period", "").lower())
            if key not in seen_metrics:
                seen_metrics.add(key)
                merged["metrics"].append(m)

        for risk in r.get("risks", []):
            key = (risk.get("name", "").lower()[:60], risk.get("category", "").lower())
            if key not in seen_risks:
                seen_risks.add(key)
                merged["risks"].append(risk)

    return merged



def generate_conclusion(risks: list[dict]) -> str:
    if not risks:
        return ""
    risk_summary = "\n".join(
        f"- {r.get('name', '?')} [{r.get('severity', '?')}]: {r.get('detailed_description', '')[:200]}"
        for r in risks
    )
    response = client.chat.completions.create(
        model=get_model(),
        temperature=0,
        messages=build_messages(CONCLUSION_PROMPT, risk_summary),
        **get_completion_options(),
    )
    return response.choices[0].message.content.strip()




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


def write_ifrs_sheet(wb, ifrs_data: dict) -> None:
    statements = ifrs_data.get("statements", [])
    if not statements:
        return

    ws = wb.create_sheet("МСФО")

    thin = Side(style="thin", color="BFBFBF")
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Collect all unique periods across all statements
    all_periods: list[str] = []
    for stmt in statements:
        for item in stmt.get("items", []):
            for p in item.get("periods", []):
                period = p.get("period", "")
                if period and period not in all_periods:
                    all_periods.append(period)

    current_row = 1

    for stmt in statements:
        stmt_type = stmt.get("type", "Форма")
        items = stmt.get("items", [])
        if not items:
            continue

        # Statement title row
        title_end_col = 2 + len(all_periods) - 1
        ws.merge_cells(start_row=current_row, start_column=1,
                       end_row=current_row, end_column=max(title_end_col, 2))
        title_cell = ws.cell(row=current_row, column=1, value=stmt_type.upper())
        title_cell.fill = PatternFill("solid", fgColor="1F4E79")
        title_cell.font = Font(bold=True, color="FFFFFF", size=12)
        title_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[current_row].height = 20
        current_row += 1

        # Header row: "Статья" + periods
        ws.cell(row=current_row, column=1, value="Статья").fill = PatternFill("solid", fgColor="BDD7EE")
        ws.cell(row=current_row, column=1).font = Font(bold=True, color="1F3864", size=10)
        ws.cell(row=current_row, column=1).border = cell_border
        ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
        for col_offset, period in enumerate(all_periods, start=2):
            cell = ws.cell(row=current_row, column=col_offset, value=period)
            cell.fill = PatternFill("solid", fgColor="BDD7EE")
            cell.font = Font(bold=True, color="1F3864", size=10)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = cell_border
        ws.row_dimensions[current_row].height = 18
        current_row += 1

        current_section = None
        for item in items:
            section = item.get("section", "")
            is_total = item.get("is_total", False)
            name = item.get("name", "")
            periods_data = {p.get("period", ""): p.get("value") for p in item.get("periods", [])}

            # Section header when section changes
            if section and section != current_section:
                current_section = section
                sec_end_col = 2 + len(all_periods) - 1
                ws.merge_cells(start_row=current_row, start_column=1,
                               end_row=current_row, end_column=max(sec_end_col, 2))
                sec_cell = ws.cell(row=current_row, column=1, value=section)
                sec_cell.fill = PatternFill("solid", fgColor="D6E4F0")
                sec_cell.font = Font(bold=True, color="1F3864", size=10)
                sec_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                sec_cell.border = cell_border
                ws.row_dimensions[current_row].height = 16
                current_row += 1

            # Data row
            name_cell = ws.cell(row=current_row, column=1, value=name)
            name_cell.alignment = Alignment(horizontal="left", vertical="center",
                                            indent=(1 if not is_total else 0), wrap_text=True)
            name_cell.border = cell_border
            if is_total:
                name_cell.font = Font(bold=True, size=10)
                name_cell.fill = PatternFill("solid", fgColor="EBF3FB")
            else:
                name_cell.font = Font(size=10)

            for col_offset, period in enumerate(all_periods, start=2):
                val = periods_data.get(period)
                val_cell = ws.cell(row=current_row, column=col_offset, value=val)
                val_cell.alignment = Alignment(horizontal="right", vertical="center")
                val_cell.border = cell_border
                if is_total:
                    val_cell.font = Font(bold=True, size=10)
                    val_cell.fill = PatternFill("solid", fgColor="EBF3FB")
                else:
                    val_cell.font = Font(size=10)

            current_row += 1

        # Spacer between statements
        current_row += 1

    # Column widths
    ws.column_dimensions["A"].width = 52
    for col_offset in range(len(all_periods)):
        col_letter = get_column_letter(2 + col_offset)
        ws.column_dimensions[col_letter].width = 18

    ws.freeze_panes = "B3"


def write_excel(data: dict, out_path: str, conclusion: str = "", ifrs_data: dict | None = None,
                company_profile: dict | None = None) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    if company_profile:
        from raexpert_context import write_company_sheets
        write_company_sheets(wb, company_profile)

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

    if ifrs_data:
        write_ifrs_sheet(wb, ifrs_data)

    # Wrap long cells (risks sheet handles alignment per cell)
    for sheet in wb.sheetnames:
        if sheet in ("Риски", "МСФО"):
            continue
        for row in wb[sheet].iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(out_path)


def main():
    parser = argparse.ArgumentParser(description="Extract key data from output.txt using the configured LLM")
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
    print(f"Split into {len(chunks)} chunk(s), sending to {get_model()}...", file=sys.stderr)

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
        f"Extracted: {len(data['metrics'])} metrics, {len(data['risks'])} risks",
        file=sys.stderr,
    )

    print("Generating conclusion...", file=sys.stderr)
    conclusion = generate_conclusion(data["risks"])

    print("Extracting IFRS financial statements...", file=sys.stderr)
    ifrs_results: list[dict] = [None] * len(chunks)
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {
            pool.submit(extract_ifrs_chunk, chunk, i + 1, len(chunks)): i
            for i, chunk in enumerate(chunks)
        }
        for future in as_completed(futures):
            idx = futures[future]
            ifrs_results[idx] = future.result()

    ifrs_data = merge_ifrs_results(ifrs_results)
    stmt_count = len(ifrs_data.get("statements", []))
    print(f"Extracted: {stmt_count} IFRS statement(s)", file=sys.stderr)

    write_excel(data, args.out, conclusion=conclusion, ifrs_data=ifrs_data)
    print(f"Saved: {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
