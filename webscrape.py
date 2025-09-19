import os
import requests
from datetime import datetime
import pandas as pd

def setup_directories(script_dir):
    input_folder = os.path.join(script_dir, "input")
    output_folder = os.path.join(script_dir, "output")
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    return input_folder, output_folder

def read_url_from_file(script_dir):
    url_file = os.path.join(script_dir, "sam_url.txt")
    try:
        with open(url_file, "r") as uf:
            url = uf.readline().strip()
        return url
    except Exception as e:
        print(f"Failed to read URL from {url_file}: {e}")
        return None

def download_csv_file(url, csv_filename):
    if os.path.exists(csv_filename):
        print(f"File for today already exists: {csv_filename}. Skipping download.")
        return
    try:
        print(f"Downloading CSV from {url} ...")
        with requests.get(url, stream=True, timeout=60) as response:
            response.raise_for_status()
            total_length = response.headers.get('content-length')
            if total_length is not None:
                total_length = int(total_length)
            downloaded = 0
            chunk_count = 0
            with open(csv_filename, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        chunk_count += 1
                        if total_length:
                            percent = (downloaded / total_length) * 100
                            print(f"Downloaded {downloaded // 1024} KB ({percent:.2f}%)", end='\r')
                        else:
                            print(f"Downloaded chunk {chunk_count}", end='\r')
            print()  # Newline after progress
        print(f"File downloaded and saved to: {csv_filename}")
    except Exception as e:
        print(f"Failed to download file: {e}")


def process_csv_file(csv_filename):
    try:
        print("Reading CSV file...")
        raw_data = pd.read_csv(csv_filename, encoding='ISO-8859-1', dtype=str, low_memory=False)
        print(f"CSV file '{csv_filename}' loaded successfully. Rows: {len(raw_data)}")

        # Save to Excel file with date-stamped filename in output directory
        print("Processing and cleaning data for Excel export...")
        today = datetime.now()
        output_dir = os.path.join(os.path.dirname(csv_filename), '..', 'output')
        output_dir = os.path.abspath(output_dir)
        os.makedirs(output_dir, exist_ok=True)
        output_filename = os.path.join(output_dir, f"{today.strftime('%Y-%m-%d')}_SamOpportunities.xlsx")
        
        # Sanitize column names and index for Excel
        print("Sanitizing column names...")
        illegal_chars = [':', '\\', '/', '?', '*', '[', ']']
        def sanitize(val):
            val = str(val)
            for ch in illegal_chars:
                val = val.replace(ch, '_')
            return val[:31]
        raw_data.columns = [sanitize(col) for col in raw_data.columns]
        
        # Remove non-printable/control characters from all string columns
        print("Cleaning string data...")
        import re
        def clean_string(s):
            if isinstance(s, str):
                # Remove all non-printable/control characters
                return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', s)
            return s
        for col in raw_data.select_dtypes(include=['object']).columns:
            raw_data[col] = raw_data[col].apply(clean_string)

        # Save to Excel with a fixed sheet name and explicit engine
        print("Saving Excel file...")
        raw_data.to_excel(output_filename, index=False, sheet_name='Sheet1', engine='openpyxl')
        print(f"Processed file saved as: {output_filename}")
    except Exception as e:
        print(f"Failed to read CSV file: {e}")

def main():
    print()
    print("SAM.gov Webscrape")
    print("=================")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder, output_folder = setup_directories(script_dir)
    url = read_url_from_file(script_dir)
    if not url:
        return
    today_str = datetime.now().strftime("%Y-%m-%d")
    csv_filename = os.path.join(input_folder, f"{today_str}_ContractOpportunitiesFullCSV.csv")
    download_csv_file(url, csv_filename)
    process_csv_file(csv_filename)

if __name__ == "__main__":
    main()
