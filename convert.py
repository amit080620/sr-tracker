#!/usr/bin/env python3
"""
SR Tracker - Excel to JSON Converter
=====================================
Run this script every morning after downloading your Excel report.
It reads the Excel file and produces data.json for the SR Tracker website.

Usage:
    python convert.py
    python convert.py myfile.xlsx
    python convert.py myfile.xlsx output.json

Requirements:
    pip install openpyxl
"""

import sys
import os
import re
import json
import glob
from datetime import datetime


# ─── CONFIG ──────────────────────────────────────────────────────────────────
# If you don't pass a filename, the script auto-finds the latest .xlsx in this folder
DEFAULT_OUTPUT = "data.json"

# Exact Excel column headers (must match your file exactly)
COL_SR          = "Case Number"
COL_NAME        = "Customer Name"
COL_DOCTOR_ID   = "DoctorId"
COL_QUERY_TYPE  = "Type of Query"
COL_SUBJECT     = "Subject"
COL_CASE_REASON = "Case Reason"
COL_ORIGIN      = "Case Origin"
COL_OPENED      = "Date/Time Opened"
COL_REOPENED    = "Date/Time Reopened"
COL_STATUS      = "Status"
COL_OWNER       = "Case Owner: Full Name"
COL_COMMENTS    = "app comments"
COL_DESCRIPTION = "Description"
# ─────────────────────────────────────────────────────────────────────────────


def clean(value):
    """Convert any cell value to a clean string."""
    if value is None:
        return ""
    s = str(value).strip()
    # Remove carriage returns that Excel sometimes adds
    s = s.replace("\r\r\n", "\n").replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("_x000D_", "")
    return s


def parse_timeline(raw_comments):
    """
    Parse app comments into structured timeline entries.
    Each entry that starts with a date pattern becomes its own item.

    Recognises patterns like:
        10/4, 18/4, 10/04, 18-04-26, 10/4/2026
    """
    if not raw_comments:
        return []

    text = clean(raw_comments)
    lines = [line.strip() for line in text.split("\n") if line.strip()]

    # Pattern: starts with DD/MM or DD/MM/YY or DD-MM-YY etc.
    date_pattern = re.compile(
        r"^(\d{1,2}[\/\-]\d{1,2}(?:[\/\-]\d{2,4})?)\s*[-–]?\s*(.*)"
    )

    entries = []
    current = None

    for line in lines:
        match = date_pattern.match(line)
        if match:
            if current:
                entries.append(current)
            date_str = match.group(1).strip()
            rest = match.group(2).strip()
            current = {"date": date_str, "text": rest}
        else:
            # Continue previous entry
            if current:
                separator = " " if current["text"] else ""
                current["text"] += separator + line
            else:
                # No date found yet, treat as a note
                current = {"date": "Note", "text": line}

    if current:
        entries.append(current)

    return entries


def find_latest_excel():
    """Auto-find the most recently modified .xlsx file in current directory."""
    files = glob.glob("*.xlsx") + glob.glob("*.xls")
    if not files:
        return None
    # Sort by modification time, newest first
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0]


def convert(excel_path, output_path):
    """Main conversion function."""
    try:
        import openpyxl
    except ImportError:
        print("ERROR: openpyxl is not installed.")
        print("Fix: pip install openpyxl")
        sys.exit(1)

    print(f"Reading: {excel_path}")

    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    except FileNotFoundError:
        print(f"ERROR: File not found: {excel_path}")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: Could not open Excel file: {e}")
        sys.exit(1)

    ws = wb.active
    print(f"Sheet: '{ws.title}'")

    # Read all rows into memory
    all_rows = list(ws.iter_rows(values_only=True))

    if not all_rows:
        print("ERROR: Excel file is empty.")
        sys.exit(1)

    # First row is headers
    headers = [str(h).strip() if h is not None else "" for h in all_rows[0]]
    data_rows = all_rows[1:]

    print(f"Headers found ({len(headers)}): {headers}")

    # Build a column index map
    col_index = {}
    for i, h in enumerate(headers):
        col_index[h] = i

    # Validate required columns
    required = [COL_SR, COL_NAME, COL_STATUS]
    missing = [c for c in required if c not in col_index]
    if missing:
        print(f"WARNING: These expected columns were not found: {missing}")
        print(f"Available columns: {headers}")

    def get_col(row, col_name):
        """Safely get a cell value by column name."""
        idx = col_index.get(col_name)
        if idx is None or idx >= len(row):
            return ""
        return clean(row[idx])

    records = []
    skipped = 0

    for i, row in enumerate(data_rows, start=2):  # start=2 because row 1 is header
        sr = get_col(row, COL_SR).strip()

        # Skip empty rows
        if not sr:
            skipped += 1
            continue

        comments_raw = get_col(row, COL_COMMENTS)
        timeline = parse_timeline(comments_raw)

        record = {
            "sr":          sr,
            "name":        get_col(row, COL_NAME),
            "doctorId":    get_col(row, COL_DOCTOR_ID),
            "queryType":   get_col(row, COL_QUERY_TYPE),
            "subject":     get_col(row, COL_SUBJECT),
            "caseReason":  get_col(row, COL_CASE_REASON),
            "origin":      get_col(row, COL_ORIGIN),
            "opened":      get_col(row, COL_OPENED),
            "reopened":    get_col(row, COL_REOPENED),
            "status":      get_col(row, COL_STATUS),
            "owner":       get_col(row, COL_OWNER),
            "comments":    comments_raw,
            "timeline":    timeline,
            "description": get_col(row, COL_DESCRIPTION),
        }
        records.append(record)

    wb.close()

    if not records:
        print("WARNING: No valid records found in the Excel file.")
        print("Check that 'Case Number' column has data.")

    # Write JSON
    output = {
        "_meta": {
            "generated":   datetime.now().strftime("%d/%m/%Y %H:%M"),
            "source_file": os.path.basename(excel_path),
            "total":       len(records),
        },
        "records": records
    }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    # Summary
    print()
    print("=" * 50)
    print(f"  SUCCESS!")
    print(f"  Records converted : {len(records)}")
    print(f"  Rows skipped      : {skipped}")
    print(f"  Output file       : {output_path}")
    print(f"  Generated at      : {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 50)

    # Status breakdown
    status_counts = {}
    for r in records:
        s = r["status"] or "Unknown"
        status_counts[s] = status_counts.get(s, 0) + 1
    print("\nStatus breakdown:")
    for s, c in sorted(status_counts.items(), key=lambda x: -x[1]):
        bar = "█" * min(c, 30)
        print(f"  {s:<30} {c:>4}  {bar}")

    print()
    print(f"Next step: Upload '{output_path}' to your GitHub Pages repo.")
    print("Done!")


def main():
    # Parse arguments
    if len(sys.argv) >= 2:
        excel_path = sys.argv[1]
    else:
        excel_path = find_latest_excel()
        if not excel_path:
            print("ERROR: No .xlsx file found in current directory.")
            print("Usage: python convert.py <filename.xlsx>")
            sys.exit(1)
        print(f"Auto-detected: {excel_path}")

    output_path = sys.argv[2] if len(sys.argv) >= 3 else DEFAULT_OUTPUT

    convert(excel_path, output_path)


if __name__ == "__main__":
    main()
