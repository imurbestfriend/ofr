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

CHUNK_SIZE = 80_000  # chars per API call


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


def extract_ifrs_chunk(chunk: str, index: int, total: int) -> dict:
    print(f"  IFRS chunk {index}/{total} ({len(chunk):,} chars)...", file=sys.stderr)
    response = client.chat.completions.create(
        model="gpt-4o",
        response_format={"type": "json_object"},
        temperature=0,
        messages=[
            {"role": "system", "content": IFRS_PROMPT},
            {"role": "user", "content": chunk},
        ],
        max_tokens=8192,
    )
    raw = response.choices[0].message.content
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        print(f"  Warning: could not parse IFRS JSON from chunk {index}", file=sys.stderr)
        return {"statements": []}


def merge_ifrs_results(results: list[dict]) -> dict:
    merged_by_type: dict[str, list] = {}
    for r in results:
        for stmt in r.get("statements", []):
            stmt_type = stmt.get("type", "").strip()
            if not stmt_type:
                continue
            if stmt_type not in merged_by_type:
                merged_by_type[stmt_type] = []
            seen = {(i.get("section", ""), i.get("name", "")) for i in merged_by_type[stmt_type]}
            for item in stmt.get("items", []):
                key = (item.get("section", ""), item.get("name", ""))
                if key not in seen:
                    merged_by_type[stmt_type].append(item)
                    seen.add(key)
    return {"statements": [{"type": t, "items": items} for t, items in merged_by_type.items()]}


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
        return {"metrics": [], "risks": []}


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
        model="gpt-4o",
        temperature=0,
        messages=[
            {"role": "system", "content": CONCLUSION_PROMPT},
            {"role": "user", "content": risk_summary},
        ],
        max_tokens=512,
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


def write_excel(data: dict, out_path: str, conclusion: str = "", ifrs_data: dict | None = None) -> None:
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
