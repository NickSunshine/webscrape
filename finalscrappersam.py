import tkinter as tk
from tkinter import filedialog
import pandas as pd
from datetime import datetime, timedelta

def process_file(file_path):
    raw_data = pd.read_csv(file_path, encoding='ISO-8859-1')
    filtered_data = raw_data[raw_data['Department/Ind.Agency'].str.contains('TRANSPORTATION, DEPARTMENT OF', na=False)]
    types_to_include = ['Presolicitation', 'Solicitation', 'Award Notice']
    filtered_data = filtered_data[filtered_data['Type'].isin(types_to_include)]
    filtered_data = filtered_data[filtered_data['BaseType'].isin(types_to_include)]
    filtered_data['PostedDate'] = pd.to_datetime(filtered_data['PostedDate'], errors='coerce', utc=True)
    filtered_data['PostedDate'] = filtered_data['PostedDate'].dt.tz_convert(None)
    
    today = datetime.now()
    two_months_ago = today - timedelta(days=60)
    filtered_data = filtered_data[(filtered_data['PostedDate'] >= two_months_ago) & (filtered_data['PostedDate'] <= today)]
    filtered_data['PostedDate'] = filtered_data['PostedDate'].dt.strftime('%m/%d/%Y')

    output_filename = f"{today.strftime('%Y-%m-%d')} Sam Opportunities.xlsx"
    filtered_data.to_excel(output_filename, index=False)
    print(f"Processed file saved as: {output_filename}")

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        process_file(file_path)

# Create Tkinter window
root = tk.Tk()
root.title("Process Contract Opportunities")

upload_button = tk.Button(root, text="Upload File", command=upload_file)
upload_button.pack(pady=20)

root.mainloop()
