import argparse
from pathlib import Path

FIELDS_PER_RECORD = 40


def preprocess_text(raw_text: str) -> list[str]:
    """Keep text EXACT, just normalize line breaks (strip blank lines)."""
    lines = raw_text.splitlines()
    cleaned: list[str] = []
    for line in lines:
        line = line.strip()
        if line:
            cleaned.append(line)
    return cleaned


def build_records(lines: list[str]) -> tuple[list[list[str]], list[str]]:
    """Build records strictly in groups of 40."""
    records: list[list[str]] = []
    current: list[str] = []
    for line in lines:
        current.append(line)
        if len(current) == FIELDS_PER_RECORD:
            records.append(current)
            current = []
    return records, current  # leftover may be half record


def merge_with_next_page(prev_half: list[str], next_lines: list[str]) -> tuple[list[list[str]], list[str]]:
    """If previous page ended with a half record, continue it with next page."""
    combined = prev_half + next_lines
    return build_records(combined)


def find_record_by_name(records: list[list[str]], target_name: str) -> list[str] | None:
    target = target_name.lower()
    for record in records:
        for field in record:
            if target in field.lower():
                return record
    return None


def print_record(record: list[str]) -> None:
    if record and len(record) == FIELDS_PER_RECORD:
        for field in record:
            print(field)
    else:
        print("Record not found or incomplete")


def write_records(records: list[list[str]], out_path: Path) -> None:
    with out_path.open("w", encoding="utf-8") as f:
        for idx, rec in enumerate(records):
            for line in rec:
                f.write(line + "\n")
            if idx != len(records) - 1:
                f.write("\n")


def main() -> int:
    parser = argparse.ArgumentParser(description="Strict 40-field grouper")
    parser.add_argument("input", nargs="?", default="output/combined_customers.txt", help="input text file")
    parser.add_argument("--target", help="name substring to print one record", default=None)
    parser.add_argument("--out", default="output/strict_customers.txt", help="output file for grouped records")
    args = parser.parse_args()

    raw_text = Path(args.input).read_text(encoding="utf-8")
    lines = preprocess_text(raw_text)

    records, leftover = build_records(lines)
    print(f"Total full records: {len(records)}")
    print(f"Incomplete fields carried forward: {len(leftover)}")

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    write_records(records, out_path)
    print(f"Grouped output written to: {out_path}")

    if args.target:
        rec = find_record_by_name(records, args.target)
        print_record(rec)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
