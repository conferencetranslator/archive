#!/usr/bin/env python3
"""Convert temp_word/*.docx transcript tables into synchronized JSON files."""

from __future__ import annotations

import json
import re
import sys
import tempfile
from pathlib import Path
from typing import Any


DIGIT_TRANSLATION = str.maketrans(
    "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
    "01234567890123456789",
)
EMPTY_PARENS = re.compile(r"\(\s*\)")
TIMESTAMP = re.compile(r"(\d{1,2}):([0-5]\d)")


def parse_timestamp(value: str, file_name: str, row_number: int) -> str:
    value = value.translate(DIGIT_TRANSLATION)
    value = re.sub(r"[\u200e\u200f\u202a-\u202e]", "", value).strip()
    match = TIMESTAMP.fullmatch(value)
    if not match:
        raise ValueError(
            f"{file_name}: invalid timestamp in data row {row_number}: {value!r}"
        )
    return f"{int(match.group(1)):02d}:{match.group(2)}"


def find_transcript_table(document: Any, file_name: str) -> Any:
    candidates = []
    for table in document.tables:
        if len(table.columns) != 3 or not table.rows:
            continue
        headers = [cell.text.strip().casefold() for cell in table.rows[0].cells]
        if (
            "زمان" in headers[0] or "time" in headers[0]
        ) and any(token in headers[1] for token in ("عربی", "عربي", "arabic")) and any(
            token in headers[2] for token in ("ترجمه", "فارسی", "persian")
        ):
            candidates.append(table)

    if len(candidates) != 1:
        raise ValueError(
            f"{file_name}: expected exactly one 3-column transcript table, "
            f"found {len(candidates)}"
        )
    return candidates[0]


def read_transcript(path: Path) -> tuple[str, list[dict[str, str]]]:
    if not path.stem.isdigit():
        raise ValueError(f"{path.name}: expected a numeric filename such as 9.docx")
    try:
        from docx import Document
    except ImportError as exc:
        raise RuntimeError(
            "python-docx is required; install it with: python -m pip install python-docx"
        ) from exc

    document = Document(path)
    table = find_transcript_table(document, path.name)
    source_rows = [row.cells for row in table.rows[1:]]
    if not source_rows:
        raise ValueError(f"{path.name}: transcript table has no data rows")

    records = []
    previous_time = "00:00"
    for row_number, cells in enumerate(source_rows, start=1):
        if len(cells) != 3:
            raise ValueError(
                f"{path.name}: data row {row_number} does not contain 3 cells"
            )
        recorded_time, arabic, persian = (cell.text.strip() for cell in cells)
        records.append({"time": previous_time, "ar": arabic, "fa": persian})
        previous_time = parse_timestamp(recorded_time, path.name, row_number)
    return path.stem, records


def insert_metadata(
    item: dict[str, Any], key: str, value: str, after_key: str
) -> None:
    if key in item:
        item[key] = value
        return
    ordered_item = {}
    inserted = False
    for existing_key, existing_value in item.items():
        ordered_item[existing_key] = existing_value
        if existing_key == after_key:
            ordered_item[key] = value
            inserted = True
    if not inserted:
        ordered_item[key] = value
    item.clear()
    item.update(ordered_item)


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    content = json.dumps(value, ensure_ascii=False, indent=2) + "\n"
    with tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8", dir=path.parent, delete=False
    ) as temporary:
        temporary.write(content)
        temporary_path = Path(temporary.name)
    try:
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)


def process_all(root: Path) -> list[tuple[str, list[dict[str, str]]]]:
    temp_dir = root / "temp_word"
    if not temp_dir.is_dir():
        raise FileNotFoundError(f"Input directory not found: {temp_dir}")
    paths = sorted(temp_dir.glob("*.docx"), key=lambda path: path.name)
    if not paths:
        print("No .docx files found in temp_word/.")
        return []

    transcripts = [read_transcript(path) for path in paths]
    data_path = root / "data.json"
    data = json.loads(data_path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("data.json must contain a JSON array")

    items_by_id = {}
    for item in data:
        if not isinstance(item, dict) or "id" not in item:
            continue
        item_id = str(item["id"])
        if item_id in items_by_id:
            raise ValueError(f"data.json contains duplicate id {item_id}")
        items_by_id[item_id] = item

    for file_id, _ in transcripts:
        if file_id not in items_by_id:
            raise ValueError(f"data.json has no item with id {file_id}")

    for file_id, records in transcripts:
        write_json(root / "texts" / f"{file_id}_sync.json", records)
        item = items_by_id[file_id]
        insert_metadata(item, "pdf_url", f"pdfs/{file_id}.pdf", "telegram_link")
        insert_metadata(item, "sync_file", f"texts/{file_id}_sync.json", "youtube_link")
    write_json(data_path, data)
    return transcripts


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    try:
        transcripts = process_all(root)
    except (OSError, ValueError, RuntimeError, json.JSONDecodeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    if transcripts:
        print("Rows below are 1-based transcript data rows (excluding the header):")
        found_empty_parentheses = False
        for file_id, records in transcripts:
            rows = [
                str(row_number)
                for row_number, record in enumerate(records, start=1)
                if EMPTY_PARENS.search(record["fa"])
            ]
            if rows:
                found_empty_parentheses = True
                print(f"{file_id}: {', '.join(rows)}")
        if not found_empty_parentheses:
            print("No empty parentheses found.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
