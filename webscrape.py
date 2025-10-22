import os
import requests
from datetime import datetime
import pandas as pd
import sys
import logging
import re

def setup_directories(script_dir):
    input_folder = os.path.join(script_dir, "input")
    output_folder = os.path.join(script_dir, "output")
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    return input_folder, output_folder

def read_url_from_file(script_dir):
    cfg_dir = os.path.join(script_dir, "cfg")
    url_file = os.path.join(cfg_dir, "sam_url.txt")
    try:
        with open(url_file, "r") as uf:
            url = uf.readline().strip()
        return url
    except Exception as e:
        logging.info(f"Failed to read URL from {url_file}: {e}")
        return None

def download_csv_file(url, csv_filename):
    if os.path.exists(csv_filename):
        logging.info(f"Contract opportunities file for today already exists: {csv_filename}. Skipping download.")
        return
    try:
        logging.info(f"Downloading CSV from {url} ...")
        with requests.get(url, stream=True, timeout=60) as response:
            response.raise_for_status()
            with open(csv_filename, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
        logging.info(f"File downloaded and saved to: {csv_filename}")
    except Exception as e:
        logging.info(f"Failed to download file: {e}")

def process_csv_file(csv_filename):
    try:
        logging.info("Reading CSV file...")
        raw_data = pd.read_csv(csv_filename, encoding='ISO-8859-1', dtype=str, low_memory=False)
        logging.info(f"CSV file loaded successfully. Rows: {len(raw_data)}, Columns: {len(raw_data.columns)}")

        # Load keywords from cfg/keywords.txt
        keywords_file = os.path.join(os.path.dirname(os.path.dirname(csv_filename)), "cfg", "keywords.txt")
        try:
            with open(keywords_file, "r", encoding="utf-8") as kf:
                keywords = [line.strip() for line in kf if line.strip()]
            logging.info(f"Loaded {len(keywords)} keywords for filtering.")
        except Exception as e:
            logging.info(f"Failed to read keywords from {keywords_file}: {e}")
            keywords = []

        if keywords:
            # Filter rows where Description contains any keyword (case-insensitive)
            pattern = '|'.join([re.escape(k) for k in keywords])
            mask = raw_data['Description'].str.contains(pattern, case=False, na=False)
            filtered_data = raw_data[mask].copy()  # <-- Add .copy() here
            logging.info(f"Filtered data: {len(filtered_data)} rows match keywords.")
        else:
            filtered_data = raw_data
            logging.info("No keywords loaded; skipping filtering.")

        # Use filtered_data for all further processing
        data_to_save = filtered_data

        # Save to Excel file with date-stamped filename in output directory
        logging.info("Processing and cleaning data for Excel export...")
        today = datetime.now()
        output_dir = os.path.join(os.path.dirname(csv_filename), '..', 'output')
        output_dir = os.path.abspath(output_dir)
        os.makedirs(output_dir, exist_ok=True)
        output_filename = os.path.join(output_dir, f"{today.strftime('%Y-%m-%d')}_SamOpportunities.xlsx")
        
        # Sanitize column names and index for Excel
        logging.info("Sanitizing column names...")
        illegal_chars = [':', '\\', '/', '?', '*', '[', ']']
        def sanitize(val):
            val = str(val)
            for ch in illegal_chars:
                val = val.replace(ch, '_')
            return val[:31]
        data_to_save.columns = [sanitize(col) for col in data_to_save.columns]
        
        # Remove non-printable/control characters from all string columns
        logging.info("Cleaning string data...")
        def clean_string(s):
            if isinstance(s, str):
                # Remove all non-printable/control characters
                return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', s)
            return s
        for col in data_to_save.select_dtypes(include=['object']).columns:
            data_to_save[col] = data_to_save[col].apply(clean_string)

        # Save to Excel with a fixed sheet name and explicit engine
        logging.info("Saving Excel file...")
        data_to_save.to_excel(output_filename, index=False, sheet_name='Sheet1', engine='openpyxl')
        logging.info(f"Processed file saved as: {output_filename}")
    except Exception as e:
        logging.info(f"Failed to read CSV file: {e}")

def main():
    # Set up logging to both file and console
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    logs_dir = os.path.join(script_dir, "logs")
    os.makedirs(logs_dir, exist_ok=True)
    log_filename = os.path.join(logs_dir, f"{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_webscrape.log")

    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[
            logging.FileHandler(log_filename, encoding="utf-8"),
            logging.StreamHandler(sys.stdout)
        ]
    )

    logging.info("SAM.gov Webscrape")
    logging.info("=================")
    input_folder, output_folder = setup_directories(script_dir)
    url = read_url_from_file(script_dir)
    if not url:
        return
    csv_filename = os.path.join(input_folder, f"ContractOpportunitiesFullCSV.csv")
    download_csv_file(url, csv_filename)
    process_csv_file(csv_filename)

if __name__ == "__main__":
    main()