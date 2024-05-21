import pandas as pd
import json
import zipfile
import os
from google.colab import files

def xlsx_to_structured_text(xlsx_file, output_text_file):
    # Read the Excel file
    xls = pd.ExcelFile(xlsx_file)

    # Dictionary to hold all sheets data
    structured_data = {}

    # Iterate through each sheet
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xlsx_file, sheet_name=sheet_name)
        # Convert each sheet to a list of dictionaries
        structured_data[sheet_name] = df.to_dict(orient='records')

    # Write the structured data to a text file in JSON format
    with open(output_text_file, 'w') as file:
        json.dump(structured_data, file, indent=4)

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
all_data = {}

for xlsx_file in xlsx_files:
    # Read the Excel file
    xls = pd.ExcelFile(xlsx_file)

    # Dictionary to hold all sheets data for this file
    structured_data = {}

    # Iterate through each sheet
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xlsx_file, sheet_name=sheet_name)
        # Convert each sheet to a list of dictionaries
        structured_data[sheet_name] = df.to_dict(orient='records')
    
    all_data[xlsx_file] = structured_data

# Write the structured data to a text file in JSON format
with open(output_text_file, 'w') as file:
    json.dump(all_data, file, indent=4)

# Download the output file
files.download(output_text_file)
