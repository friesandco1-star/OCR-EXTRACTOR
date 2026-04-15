#!/usr/bin/env python3
"""
Strict OCR-fragile customer extractor and formatter.

Rules implemented:
- Parse visual order top-to-bottom.
- Record starts on 6-digit line.
- Keep exact 40-field sequence.
- Carry unfinished bottom record into next page top.
- If unreadable/missing, write token (default: 0).
- Output required "Customer N" + numbered field labels.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

FIELDS = [
    "Record No.",
    "Customer Name",
    "Relative Name",
    "CustomerAddress",
    "City1",
    "ZipCode1",
    "State1",
    "Contact Number",
    "Alternative No.",
    "Date Of Birth",
    "Representa. Name",
    "Referral Date",
    "Invoice No.",
    "Lot Ref",
    "Branch Ref",
    "CHeadNo",
    "Customer Email Id",
    "Customer Ref No.",
    "NomineeName",
    "Nominee EmailId",
    "Street Address 2",
    "City2",
    "State2",
    "ZipCode2",
    "Total Net Worth",
    "Monthly Income",
    "Medicare",
    "LoanAmount",
    "Total Debts",
    "Insu Policy No.",
    "BatchNum",
    "SumAssured",
    "PremiumAmount",
    "MaturityAmount",
    "AgencyCode",
    "AgentName",
    "AgentCode",
    "AgentEmail",
    "AgentPhone",
    "Credit Card No.",
]

FIELDS_PER_RECORD = 40

TITLE_RE = re.compile(r"^(?:mr|mrs|ms|dr)\.", re.I)
RECORD_RE = re.compile(r"^\d{6}\b")
PHONE_RE = re.compile(r"^\(\d{3}\)\s\d{3}-\d{4}$")
DATE_RE = re.compile(
    r"^(?:"
    r"[A-Za-z]+,\s[A-Za-z]+\s\d{2},\s\d{4}|"
    r"\d{1,2}/[A-Za-z]+/\d{2,4}|"
    r"\d{1,2}-[A-Za-z]{3}-\d{2,4}"
    r")$"
)
EMAIL_RE = re.compile(r"^[A-Za-z0-9_.+!|/-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")


def normalize_lines(raw: str) -> list[str]:
    lines = []
    for line in raw.splitlines():
        line = line.strip()
        if line:
            lines.append(line)
    return lines


def split_pages(lines: list[str]) -> list[list[str]]:
    pages: list[list[str]] = []
    current: list[str] = []
    for line in lines:
        if set(line) == {"="}:
            if current:
                pages.append(current)
                current = []
            continue
        if line.lower().endswith((".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp")):
            continue
        current.append(line)
    if current:
        pages.append(current)
    return pages


def split_record_chunks(page_lines: list[str]) -> tuple[list[str], list[list[str]]]:
    leading: list[str] = []
    chunks: list[list[str]] = []
    current: list[str] = []
    started = False

    for line in page_lines:
        if RECORD_RE.match(line):
            if current:
                chunks.append(current)
            current = [line]
            started = True
        else:
            if not started:
                leading.append(line)
            else:
                current.append(line)

    if current:
        chunks.append(current)
    return leading, chunks


def is_numeric_text(value: str) -> bool:
    return bool(re.fullmatch(r"[0-9][0-9 \-]*[0-9]|\d+", value))


def field_match(field_idx: int, line: str) -> bool:
    if not line:
        return False

    if field_idx == 0:
        return bool(RECORD_RE.match(line))
    if field_idx in {1, 2, 10, 18, 35}:
        return bool(TITLE_RE.match(line))
    if field_idx == 3:
        return bool(re.match(r"^[A-Za-z0-9]", line))
    if field_idx == 4:
        return bool(re.match(r"^[A-Za-z]", line))
    if field_idx == 5:
        return is_numeric_text(line)
    if field_idx == 6:
        return bool(re.match(r"^[A-Za-z]", line))
    if field_idx in {7, 8}:
        return bool(PHONE_RE.fullmatch(line))
    if field_idx in {9, 11}:
        return bool(DATE_RE.fullmatch(line))
    if field_idx == 12:
        return bool(re.match(r"^[A-Za-z]", line))
    if field_idx == 13:
        return bool(re.match(r"^[A-Za-z]", line))
    if field_idx == 14:
        return bool(re.match(r"^[A-Za-z0-9]", line))
    if field_idx == 15:
        return is_numeric_text(line)
    if field_idx in {16, 19, 37}:
        return bool(EMAIL_RE.fullmatch(line))
    if field_idx == 17:
        return bool(re.match(r"^[A-Za-z]", line))
    if field_idx == 20:
        return bool(re.match(r"^[A-Za-z0-9]", line))
    if field_idx in {21, 22}:
        return bool(re.match(r"^[A-Za-z]", line))
    if field_idx == 23:
        return is_numeric_text(line) or bool(re.fullmatch(r"[A-Za-z0-9]{3,10}", line.replace(" ", "")))
    if field_idx in {24, 25, 26, 27, 28, 30, 31, 32, 33}:
        return is_numeric_text(line)
    if field_idx == 29:
        return bool(re.match(r"^[A-Za-z]", line))
    if field_idx == 34:
        return bool(re.match(r"^[A-Za-z]", line))
    if field_idx == 36:
        return bool(re.match(r"^\d+\\", line))
    if field_idx == 38:
        return bool(re.search(r"\d", line))
    if field_idx == 39:
        return bool(re.search(r"(card|visa|mastercard|discover|american express|diners)", line, re.I)) or bool(
            re.search(r"\d{10,}", line)
        )
    return True


def parse_chunk_to_fields(chunk: list[str], fill_token: str) -> list[str]:
    fields = [fill_token] * FIELDS_PER_RECORD
    line_idx = 0
    field_idx = 0

    flexible = {3, 4, 6, 12, 13, 14, 20, 21, 22, 29, 34, 38, 39}

    while field_idx < FIELDS_PER_RECORD and line_idx < len(chunk):
        line = chunk[line_idx]
        if field_match(field_idx, line):
            fields[field_idx] = line
            line_idx += 1
            field_idx += 1
            continue

        # small lookahead; if a near line matches expected field, mark current missing
        found_ahead = False
        for step in range(1, 4):
            pos = line_idx + step
            if pos < len(chunk) and field_match(field_idx, chunk[pos]):
                fields[field_idx] = fill_token
                field_idx += 1
                found_ahead = True
                break
        if found_ahead:
            continue

        # consume uncertain text only for flexible fields
        if field_idx in flexible:
            fields[field_idx] = line
            line_idx += 1
            field_idx += 1
        else:
            fields[field_idx] = fill_token
            field_idx += 1

    # If there are leftover OCR lines and credit-card field is still empty, use remaining text there.
    if line_idx < len(chunk) and (fields[39] == fill_token):
        tail = " ".join(chunk[line_idx:]).strip()
        if tail:
            fields[39] = tail

    return fields


def extract_records(lines: list[str], fill_token: str, partial_tag: str) -> list[list[str]]:
    pages = split_pages(lines)
    records: list[list[str]] = []
    carry_chunk: list[str] = []

    for page in pages:
        leading, chunks = split_record_chunks(page)

        if carry_chunk:
            carry_chunk.extend(leading)
            if chunks:
                records.append(parse_chunk_to_fields(carry_chunk, fill_token))
                carry_chunk = []
            else:
                continue
        elif leading:
            # no active carry; ignore leading noise
            pass

        if not chunks:
            continue

        for i, chunk in enumerate(chunks):
            if i == len(chunks) - 1:
                # hold last chunk for possible continuation with next page top
                carry_chunk = chunk[:]
            else:
                records.append(parse_chunk_to_fields(chunk, fill_token))

    # finalize leftover last chunk
    if carry_chunk:
        final = parse_chunk_to_fields(carry_chunk, fill_token)
        if final[0] == fill_token:
            final[0] = partial_tag
        elif partial_tag not in final[0]:
            final[0] = f"{final[0]} {partial_tag}"
        records.append(final)

    return records


def format_output(records: list[list[str]]) -> str:
    out: list[str] = []
    for idx, rec in enumerate(records, start=1):
        out.append(f"Customer {idx}")
        for i, label in enumerate(FIELDS, start=1):
            out.append(f"{i}. {label}: {rec[i - 1]}")
        out.append("")
    return "\n".join(out).rstrip() + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Strict 40-sequence OCR customer formatter")
    parser.add_argument("input", nargs="?", default="/Users/saru/Documents/input.txt", help="raw OCR input file")
    parser.add_argument("--out", default="/Users/saru/Documents/strict_output.txt", help="formatted output file")
    parser.add_argument("--fill-token", default="0", help="token for missing/unreadable fields (default: 0)")
    parser.add_argument("--partial-tag", default="[partial record]", help="tag appended to cut-off records")
    args = parser.parse_args()

    raw = Path(args.input).read_text(encoding="utf-8", errors="ignore")
    lines = normalize_lines(raw)
    records = extract_records(lines, fill_token=args.fill_token, partial_tag=args.partial_tag)
    formatted = format_output(records)

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(formatted, encoding="utf-8")
    print(f"Wrote {len(records)} customers to {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
