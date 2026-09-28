#!/usr/bin/env python3
import base64
import io
import subprocess
import sys
from pathlib import Path

from openai import OpenAI
from pdf2image import convert_from_path
from dotenv import load_dotenv

load_dotenv()

client = OpenAI()

MAX_CHARS = 400_000

OCR_PROMPT = (
    "Извлеки весь текст с этой страницы финансового отчёта. "
    "Таблицы оформляй в формате Markdown (| заголовок | заголовок |\n|---|---| и т.д.). "
    "Числа передавай точно как в документе, включая апострофы как разделители тысяч "
    "и скобки для отрицательных значений. "
    "Выводи только извлечённый текст, без пояснений."
)


def write_output(text: str, output_path: str) -> None:
    chunks = text.split("\n\n---\n\n")
    parts, current, current_size = [], [], 0
    for chunk in chunks:
        size = len(chunk) + (8 if current else 0)
        if current and current_size + size > MAX_CHARS:
            parts.append("\n\n---\n\n".join(current))
            current, current_size = [chunk], len(chunk)
        else:
            current.append(chunk)
            current_size += size
    if current:
        parts.append("\n\n---\n\n".join(current))

    base = Path(output_path)
    if len(parts) == 1:
        base.write_text(parts[0], encoding="utf-8")
        print(f"Saved: {base} ({len(parts[0]):,} chars)", file=sys.stderr)
    else:
        for i, part in enumerate(parts, start=1):
            out = base.parent / f"{base.stem}_part{i}{base.suffix}"
            out.write_text(part, encoding="utf-8")
            print(f"Saved: {out} ({len(part):,} chars)", file=sys.stderr)


def get_page_count(pdf_path: str) -> int:
    result = subprocess.run(["pdfinfo", pdf_path], capture_output=True, text=True)
    for line in result.stdout.splitlines():
        if line.startswith("Pages:"):
            return int(line.split(":")[1].strip())
    raise RuntimeError(f"Could not determine page count for {pdf_path}")


def extract_embedded_text(pdf_path: str, page_num: int) -> str:
    result = subprocess.run(
        ["pdftotext", "-layout", "-f", str(page_num), "-l", str(page_num), pdf_path, "-"],
        capture_output=True,
        text=True,
    )
    return result.stdout.strip()


def ocr_with_gpt(pdf_path: str, page_num: int) -> str:
    images = convert_from_path(pdf_path, dpi=200, first_page=page_num, last_page=page_num)
    buf = io.BytesIO()
    images[0].save(buf, format="PNG")
    image_b64 = base64.b64encode(buf.getvalue()).decode()

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": OCR_PROMPT},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{image_b64}",
                            "detail": "high",
                        },
                    },
                ],
            }
        ],
        max_tokens=4096,
    )
    text = response.choices[0].message.content
    # Strip markdown code fences GPT sometimes wraps output in
    if text.startswith("```"):
        text = text.split("\n", 1)[-1]
    if text.endswith("```"):
        text = text.rsplit("```", 1)[0]
    return text.strip()


def extract_pdf(pdf_path: str) -> str:
    pdf_path = str(Path(pdf_path).resolve())
    page_count = get_page_count(pdf_path)

    pages = []
    for page_num in range(1, page_count + 1):
        text = extract_embedded_text(pdf_path, page_num)
        if text:
            print(f"  [{page_num}/{page_count}] embedded text ({len(text)} chars)", file=sys.stderr)
            pages.append(text)
        else:
            print(f"  [{page_num}/{page_count}] scanned page — calling GPT-4o...", file=sys.stderr)
            text = ocr_with_gpt(pdf_path, page_num)
            print(f"  [{page_num}/{page_count}] OCR done ({len(text)} chars)", file=sys.stderr)
            pages.append(text)

    return "\n\n---\n\n".join(pages)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("pdfs", nargs="+", metavar="PDF")
    parser.add_argument("-o", "--output", metavar="FILE",
                        help="Output file. Auto-splits into _part1, _part2… if output exceeds MAX_CHARS. Defaults to stdout.")
    args = parser.parse_args()

    for path in args.pdfs:
        if not Path(path).exists():
            print(f"File not found: {path}", file=sys.stderr)
            sys.exit(1)

    parts = []
    for i, path in enumerate(args.pdfs, start=1):
        print(f"Processing ({i}/{len(args.pdfs)}): {path}", file=sys.stderr)
        text = extract_pdf(path)
        header = f"=== ДОКУМЕНТ {i}: {Path(path).name} ==="
        parts.append(f"{header}\n\n{text}")

    result = "\n\n---\n\n".join(parts)

    if args.output:
        write_output(result, args.output)
    else:
        print(result)
