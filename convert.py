import pandas as pd
import json
import os
import re


def clean_text(val):
    if pd.isna(val):
        return ""
    return (
        str(val)
        .replace("_x000D_", "")
        .replace("\r\n", " ")
        .replace("\r", " ")
        .replace("\n", " ")
        .strip()
    )


def split_timeline(val):
    if pd.isna(val):
        return []

    text = (
        str(val)
        .replace("_x000D_", "\n")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
    )

    lines = [x.strip() for x in text.split("\n") if x.strip()]

    # Matches: 10/4, 18/4, 10/04, 18-04-26, 10/4/2026
    date_pattern = re.compile(
        r"^(\d{1,2}[\/\-]\d{1,2}(?:[\/\-]\d{2,4})?)\s*[-–]?\s*(.*)"
    )

    timeline = []
    current = None

    for line in lines:
        match = date_pattern.match(line)
        if match:
            if current:
                timeline.append(current)
            current = {
                "date": match.group(1).strip(),
                "text": match.group(2).strip()
            }
        else:
            if current:
                current["text"] += " " + line
            else:
                current = {"date": "", "text": line}

    if current:
        timeline.append(current)

    return timeline


def convert():
    file_path = "data.xlsx"

    if not os.path.exists(file_path):
        print("ERROR: data.xlsx not found. Put your Excel file in the same folder.")
        return

    df = pd.read_excel(file_path)
    df = df.dropna(how="all")

    data = []

    for _, row in df.iterrows():
        sr = clean_text(row.get("Case Number"))
        if not sr:
            continue

        data.append({
            "sr":       sr,
            "name":     clean_text(row.get("Customer Name")),
            "status":   clean_text(row.get("Status")),
            "owner":    clean_text(row.get("Case Owner: Full Name", row.get("Case Owner", ""))),
            "opened":   clean_text(row.get("Date/Time Opened")),
            "subject":  clean_text(row.get("Subject")),
            "origin":   clean_text(row.get("Case Origin")),
            "doctorId": clean_text(row.get("DoctorId")),
            "reopened": clean_text(row.get("Date/Time Reopened")),
            "timeline": split_timeline(row.get("app comments"))
        })

    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"Done! {len(data)} records saved to data.json")


if __name__ == "__main__":
    convert()
