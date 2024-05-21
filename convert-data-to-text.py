import pandas as pd
import json
import zipfile
import os
from google.colab import files

def convert_to_serializable(obj):
    if isinstance(obj, (pd.Timestamp, pd._libs.tslibs.timestamps.Timestamp)):
        return obj.isoformat()
    elif isinstance(obj, pd.Series):
        return obj.tolist()
    elif isinstance(obj, pd.DataFrame):
        return obj.to_dict()
    else:
        return str(obj)

def xlsx_to_structured_text(xlsx_file, output_text_file):
    # Read the Excel file
    xls = pd.ExcelFile(xlsx_file)

    # List to hold all sheets data
    structured_data = []

    # Iterate through each sheet
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xlsx_file, sheet_name=sheet_name)
        # Convert each sheet to a list of dictionaries
        sheet_data = df.applymap(convert_to_serializable).to_dict(orient='records')
        structured_data.append({sheet_name: sheet_data})

    # Write the structured data to a text file in a simple key-value format
    with open(output_text_file, 'w') as file:
        for sheet in structured_data:
            for sheet_name, data in sheet.items():
                file.write(f"Sheet: {sheet_name}\n")
                for record in data:
                    file.write(f"{json.dumps(record)}\n")
                file.write("\n")

# Upload XLSX or ZIP file
uploaded = files.upload()

for file_name in uploaded.keys():
    if file_name.endswith('.zip'):
        with zipfile.ZipFile(file_name, 'r') as zip_ref:
            zip_ref.extractall()
        xlsx_files = [f for f in os.listdir() if f.endswith('.xlsx')]
    elif file_name.endswith('.xlsx'):
        xlsx_files = [file_name]
    else:
        xlsx_files = []

output_text_file = 'output_file.txt'
all_data = []

for xlsx_file in xlsx_files:
    xlsx_to_structured_text(xlsx_file, output_text_file)

# Download the output file
files.download(output_text_file)
