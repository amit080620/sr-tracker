import pandas as pd
import json
import os

def convert():
    file_path = 'data.xlsx'
    if not os.path.exists(file_path):
        print("Waiting for data.xlsx...")
        return

    # Read the excel you will upload
    df = pd.read_excel(file_path)
    
    # Matches the columns from your 'Filed ai feed' file
    data = []
    for _, row in df.iterrows():
        data.append({
            "sr": str(row.get('Case Number', '')),
            "name": str(row.get('Customer Name', '')),
            "status": str(row.get('Status', '')),
            "owner": str(row.get('Case Owner: Full Name', row.get('Case Owner', ''))),
            "opened": str(row.get('Date/Time Opened', '')),
            "subject": str(row.get('Subject', '')),
            "comments": str(row.get('app comments', ''))
        })

    with open('data.json', 'w') as f:
        json.dump(data, f, indent=2)

if __name__ == "__main__":
    convert()
