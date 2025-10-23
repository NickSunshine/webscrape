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

def process_csv_file(csv_filename, timestamp):
    try:
        logging.info("\nReading CSV file...")
        raw_data = pd.read_csv(csv_filename, encoding='ISO-8859-1', dtype=str, low_memory=False)
        logging.info(f"CSV file loaded successfully. Rows: {len(raw_data)}, Columns: {len(raw_data.columns)}")

        keywords_file = os.path.join(os.path.dirname(os.path.dirname(csv_filename)), "cfg", "keywords.txt")
        try:
            with open(keywords_file, "r", encoding="utf-8") as kf:
                keywords = [line.strip() for line in kf if line.strip()]
            logging.info(f"\nLoaded {len(keywords)} keywords for filtering from keywords.txt.")
        except Exception as e:
            logging.info(f"Failed to read keywords from {keywords_file}: {e}")
            keywords = []

        # Define pattern/flags function for both filtering and counting
        def make_pattern_and_flags(k):
            if " " in k:
                # Phrase: always case-insensitive
                pattern = r'(?<!\w)' + re.escape(k) + r'(?!\w)'
                flags = re.IGNORECASE
            elif k.isalpha() and k.isupper():
                # All-uppercase single word: case-sensitive
                pattern = r'\b' + re.escape(k) + r'\b'
                flags = 0
            else:
                # Other single word: case-insensitive
                pattern = r'\b' + re.escape(k) + r'\b'
                flags = re.IGNORECASE
            return pattern, flags

        if keywords:
            # Build a mask for each keyword, then combine
            masks = []
            for k in keywords:
                pattern, flags = make_pattern_and_flags(k)
                logging.info(f"Searching for keyword: '{k}' ...")
                mask = raw_data['Description'].str.contains(pattern, case=False, na=False, regex=True, flags=flags)
                match_count = mask.sum()
                logging.info(f"Found {match_count} row(s) matching keyword: '{k}'")
                masks.append(mask)
            if masks:
                combined_mask = masks[0]
                for m in masks[1:]:
                    combined_mask = combined_mask | m
                filtered_data = raw_data[combined_mask].copy()
                logging.info(f"\n{len(filtered_data)} rows match keywords.")
            else:
                filtered_data = raw_data
                logging.info("No keywords loaded; skipping filtering.")
        else:
            filtered_data = raw_data
            logging.info("No keywords loaded; skipping filtering.")

        # Filter columns based on keep_cols.txt
        keep_cols_file = os.path.join(os.path.dirname(os.path.dirname(csv_filename)), "cfg", "keep_cols.txt")
        try:
            with open(keep_cols_file, "r", encoding="utf-8") as kcf:
                keep_cols = [line.strip() for line in kcf if line.strip()]
            # Only keep columns that exist in the DataFrame
            cols_to_keep = [col for col in keep_cols if col in filtered_data.columns]
            filtered_data = filtered_data[cols_to_keep]
            logging.info(f"Keeping {len(cols_to_keep)} columns as specified in keep_cols.txt.")
        except Exception as e:
            logging.info(f"Failed to read or apply keep_cols.txt: {e}")

        # Save to Excel file with date-stamped filename in output directory
        logging.info("\nProcessing and cleaning data for Excel export...")
        output_dir = os.path.join(os.path.dirname(csv_filename), '..', 'output')
        output_dir = os.path.abspath(output_dir)
        os.makedirs(output_dir, exist_ok=True)
        output_filename = os.path.join(output_dir, f"{timestamp}_SAMOpportunities.xlsx")
        
        # Sanitize column names and index for Excel
        logging.info("Sanitizing column names...")
        illegal_chars = [':', '\\', '/', '?', '*', '[', ']']
        def sanitize(val):
            val = str(val)
            for ch in illegal_chars:
                val = val.replace(ch, '_')
            return val[:31]
        filtered_data.columns = [sanitize(col) for col in filtered_data.columns]
        
        # Remove non-printable/control characters from all string columns
        logging.info("Cleaning string data...")
        def clean_string(s):
            if isinstance(s, str):
                # Remove all non-printable/control characters
                return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', s)
            return s
        for col in filtered_data.select_dtypes(include=['object']).columns:
            filtered_data.loc[:, col] = filtered_data[col].apply(clean_string)

        # Save to Excel with a renamed sheet and add keyword match counts sheet
        logging.info("Saving Excel file...")

        # Prepare keyword match counts
        keyword_counts = []
        if keywords:
            for k in keywords:
                pattern, flags = make_pattern_and_flags(k)
                count = raw_data['Description'].str.count(pattern, flags=flags).sum()
                keyword_counts.append({'Keyword': k, 'Count': int(count)})
            keyword_counts_df = pd.DataFrame(keyword_counts)
            keyword_counts_df = keyword_counts_df.sort_values(by='Count', ascending=False).reset_index(drop=True)
        else:
            keyword_counts_df = pd.DataFrame(columns=['Keyword', 'Count'])

        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            filtered_data.to_excel(writer, index=False, sheet_name='Keyword Matches')
            keyword_counts_df.to_excel(writer, index=False, sheet_name='Keyword Match Counts')
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
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    log_filename = os.path.join(logs_dir, f"{timestamp}_webscrape.log")

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
    process_csv_file(csv_filename, timestamp)

if __name__ == "__main__":
    main()