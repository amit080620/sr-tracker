import pandas as pd
import json
import os

def clean_text(val):
    if pd.isna(val):
        return ""
    return str(val).replace("_x000D_", "").replace("\n", " ").strip()

def split_comments(val):
    if pd.isna(val):
        return []
    text = str(val).replace("_x000D_", "\n")
    return [x.strip() for x in text.split("\n") if x.strip()]

def convert():
    file_path = 'data.xlsx'
    if not os.path.exists(file_path):
        return

    df = pd.read_excel(file_path)

    data = []
    for _, row in df.iterrows():
        data.append({
            "sr": clean_text(row.get('Case Number')),
            "name": clean_text(row.get('Customer Name')),
            "status": clean_text(row.get('Status')),
            "owner": clean_text(row.get('Case Owner: Full Name', row.get('Case Owner'))),
            "opened": clean_text(row.get('Date/Time Opened')),
            "subject": clean_text(row.get('Subject')),
            "origin": clean_text(row.get('Origin')),
            "doctor_id": clean_text(row.get('Doctor ID')),
            "reopened": clean_text(row.get('Reopened Date')),
            "comments": split_comments(row.get('app comments'))
        })

    with open('data.json', 'w') as f:
        json.dump(data, f, indent=2)

if __name__ == "__main__":
    convert()
