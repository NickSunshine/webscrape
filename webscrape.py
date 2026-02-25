
# Standard library imports
import os
import sys
import re
import logging
from datetime import datetime

# Third-party imports
import requests      # For HTTP requests to download CSV files
import pandas as pd  # For data manipulation and Excel export
from openpyxl.styles import Alignment  # For Excel formatting
# tzlocal is used later for timezone handling (see below)

def setup_directories(script_dir):
    """
    Ensure the input and output directories exist within the given script directory.

    Args:
        script_dir (str): The absolute path to the directory containing the script or executable.

    Returns:
        tuple: (input_folder, output_folder) - Absolute paths to the input and output directories.

    Notes:
        - Creates the directories if they do not already exist.
        - Used to organize input CSVs and output Excel files for the webscrape process.
    """
    input_folder = os.path.join(script_dir, "input")
    output_folder = os.path.join(script_dir, "output")
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    return input_folder, output_folder

def read_url_from_file(script_dir):
    """
    Reads the download URL from the configuration file 'sam_url.txt' in the cfg directory.

    Args:
        script_dir (str): The absolute path to the directory containing the script or executable.

    Returns:
        str or None: The URL string if successfully read, otherwise None.

    Notes:
        - Expects the file 'cfg/sam_url.txt' to exist and contain the URL on the first line.
        - Logs an info message and returns None if the file cannot be read.
    """
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
    """
    Downloads a CSV file from the specified URL and saves it to the given filename.

    Args:
        url (str): The URL to download the CSV file from.
        csv_filename (str): The local file path where the downloaded CSV will be saved.

    Returns:
        None

    Notes:
        - Logs progress and errors using the logging module.
        - Uses streaming download to handle large files efficiently.
        - Overwrites the file if it already exists.
    """
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
    # ...existing code...
    try:
        logging.info("\nReading CSV file...")
        raw_data = pd.read_csv(csv_filename, encoding='ISO-8859-1', dtype=str, low_memory=False)
        logging.info(f"CSV file loaded successfully. Rows: {len(raw_data)}, Columns: {len(raw_data.columns)}")


        # --- Organization filtering (keyorgs.txt on 'Sub-Tier') ---
        keyorgs_file = os.path.join(os.path.dirname(os.path.dirname(csv_filename)), "cfg", "keyorgs.txt")
        try:
            with open(keyorgs_file, "r", encoding="utf-8") as kof:
                keyorgs = [line.strip() for line in kof if line.strip()]
            logging.info(f"Loaded {len(keyorgs)} organization filters from keyorgs.txt.")
        except Exception as e:
            logging.info(f"Failed to read keyorgs from {keyorgs_file}: {e}")
            keyorgs = []

        def make_pattern_and_flags(k):
            if " " in k:
                pattern = r'(?<!\w)' + re.escape(k) + r'(?!\w)'
                flags = re.IGNORECASE
            elif k.isalpha() and k.isupper():
                pattern = r'\b' + re.escape(k) + r'\b'
                flags = 0
            else:
                pattern = r'\b' + re.escape(k) + r'\b'
                flags = re.IGNORECASE
            return pattern, flags

        # Filter by organizations first (on 'Sub-Tier')
        org_filtered_data = raw_data
        if keyorgs:
            org_masks = []
            for org in keyorgs:
                pattern, flags = make_pattern_and_flags(org)
                mask = org_filtered_data['Sub-Tier'].str.contains(pattern, case=False, na=False, regex=True, flags=flags)
                org_masks.append(mask)
            if org_masks:
                combined_org_mask = org_masks[0]
                for m in org_masks[1:]:
                    combined_org_mask = combined_org_mask | m
                org_filtered_data = org_filtered_data[combined_org_mask].copy()
                logging.info(f"{len(org_filtered_data)} rows match organization filters.")
            else:
                logging.info("No organization filters loaded; skipping org filtering.")
        else:
            logging.info("No organization filters loaded; skipping org filtering.")

        # --- Keyword filtering (keywords.txt on 'Description') ---
        keywords_file = os.path.join(os.path.dirname(os.path.dirname(csv_filename)), "cfg", "keywords.txt")
        try:
            with open(keywords_file, "r", encoding="utf-8") as kf:
                keywords = [line.strip() for line in kf if line.strip()]
            logging.info(f"\nLoaded {len(keywords)} keywords for filtering from keywords.txt.")
        except Exception as e:
            logging.info(f"Failed to read keywords from {keywords_file}: {e}")
            keywords = []

        if keywords:
            # Use index mapping to avoid list index out of range
            org_indices = list(org_filtered_data.index)
            matched_keywords_per_row = {idx: [] for idx in org_indices}
            for k in keywords:
                pattern, flags = make_pattern_and_flags(k)
                logging.info(f"Searching for keyword: '{k}' ...")
                mask = org_filtered_data['Description'].str.contains(pattern, case=False, na=False, regex=True, flags=flags)
                for idx, is_match in zip(org_indices, mask):
                    if is_match:
                        matched_keywords_per_row[idx].append(k)
            # Only keep rows where at least one keyword matched
            matched_mask = org_filtered_data.index.map(lambda idx: len(matched_keywords_per_row[idx]) > 0)
            filtered_data = org_filtered_data[matched_mask].copy()
            logging.info(f"\n{len(filtered_data)} rows match keywords (at least one keyword).")
            # Build the columns for only the filtered rows
            filtered_indices = filtered_data.index.tolist()
            matched_keywords_filtered = [", ".join(matched_keywords_per_row[i]) for i in filtered_indices]
            matched_keywords_count = [len(matched_keywords_per_row[i]) for i in filtered_indices]
            filtered_data.insert(0, "Matched Keywords", matched_keywords_filtered)
            filtered_data.insert(1, "Matched Keywords Count", matched_keywords_count)
            filtered_data = filtered_data.sort_values(
                by=["Matched Keywords Count", "Matched Keywords"],
                ascending=[False, True]
            ).reset_index(drop=True)
        else:
            filtered_data = org_filtered_data
            logging.info("No keywords loaded; skipping filtering.")
            filtered_data.insert(0, "Matched Keywords", [""] * len(filtered_data))

        # Filter columns based on keep_cols.txt, but always include "Matched Keywords" and "Matched Keywords Count" as the first columns
        keep_cols_file = os.path.join(os.path.dirname(os.path.dirname(csv_filename)), "cfg", "keep_cols.txt")
        try:
            with open(keep_cols_file, "r", encoding="utf-8") as kcf:
                keep_cols = [line.strip() for line in kcf if line.strip()]
            # Only keep columns that exist in the DataFrame
            cols_to_keep = [col for col in keep_cols if col in filtered_data.columns]
            # Ensure "Matched Keywords" and "Matched Keywords Count" are the first columns
            for special_col in ["Matched Keywords", "Matched Keywords Count"]:
                if special_col not in cols_to_keep:
                    cols_to_keep = [special_col] + cols_to_keep
                else:
                    # Move to the front if already present
                    cols_to_keep = [special_col] + [col for col in cols_to_keep if col != special_col]
            filtered_data = filtered_data[cols_to_keep]
            logging.info(f"Keeping {len(cols_to_keep)} columns as specified in keep_cols.txt (plus Matched Keywords columns).")
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

        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            filtered_data.to_excel(writer, index=False, sheet_name='Keyword Matches')

            # --- Keyword Match Counts sheet ---
            keyword_counts = []
            if keywords:
                for k in keywords:
                    pattern, flags = make_pattern_and_flags(k)
                    mask = filtered_data['Description'].str.contains(pattern, case=False, na=False, regex=True, flags=flags)
                    count = mask.sum()
                    keyword_counts.append({'Keyword': k, 'Count': int(count)})
                keyword_counts_df = pd.DataFrame(keyword_counts)
                keyword_counts_df = keyword_counts_df.sort_values(by='Count', ascending=False).reset_index(drop=True)
            else:
                keyword_counts_df = pd.DataFrame(columns=['Keyword', 'Count'])
            keyword_counts_df.to_excel(writer, index=False, sheet_name='Keyword Match Counts')

            # Format both sheets
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for cell in next(worksheet.iter_rows(min_row=1, max_row=1)):
                    cell.alignment = Alignment(horizontal='left')
                for col in worksheet.columns:
                    max_length = 0
                    col_letter = col[0].column_letter
                    for cell in col[1:]:
                        try:
                            cell_length = len(str(cell.value)) if cell.value is not None else 0
                            if cell_length > max_length:
                                max_length = cell_length
                        except:
                            pass
                    header_length = len(str(col[0].value)) if col[0].value is not None else 0
                    best_length = max(max_length, header_length) + 2
                    worksheet.column_dimensions[col_letter].width = best_length

                for col_idx, cell in enumerate(next(worksheet.iter_rows(min_row=1, max_row=1)), 1):
                    if cell.value == "Link":
                        for row in worksheet.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                            link_cell = row[0]
                            if link_cell.value and isinstance(link_cell.value, str) and link_cell.value.startswith("http"):
                                link_cell.hyperlink = link_cell.value
                                link_cell.style = "Hyperlink"
                        break

        # --- Highlight new rows in yellow based on PostedDate compared to previous Excel file timestamp ---
        try:
            from openpyxl import load_workbook
            from openpyxl.styles import PatternFill
            import re as _re
            from dateutil.parser import parse as dtparse
            import tzlocal
            local_tz = tzlocal.get_localzone()
            output_files = [f for f in os.listdir(output_dir) if f.endswith('_SAMOpportunities.xlsx')]
            output_files = sorted(output_files)
            if len(output_files) >= 2:
                prev_file = output_files[-2]
                match = _re.match(r"(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})_SAMOpportunities\.xlsx$", prev_file)
                if match:
                    prev_timestamp_str = match.group(1)
                    prev_timestamp_naive = datetime.strptime(prev_timestamp_str, "%Y-%m-%d_%H-%M-%S")
                    # Make filename timestamp offset-aware (system local time)
                    prev_timestamp = prev_timestamp_naive.replace(tzinfo=local_tz)
                    new_row_indices = []
                    parse_failures = 0
                    failed_examples = []
                    for idx, val in enumerate(filtered_data["PostedDate"].fillna("").astype(str)):
                        date_str = val.strip()
                        try:
                            # Try dateutil first
                            row_dt = dtparse(date_str, fuzzy=True)
                        except Exception:
                            try:
                                row_dt = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S.%f-%z")
                            except Exception:
                                parse_failures += 1
                                if len(failed_examples) < 5:
                                    failed_examples.append(date_str)
                                continue
                        # Convert row_dt to system local time for comparison
                        if row_dt.tzinfo is not None:
                            row_dt_local = row_dt.astimezone(local_tz)
                        else:
                            row_dt_local = row_dt.replace(tzinfo=local_tz)
                        if row_dt_local > prev_timestamp:
                            new_row_indices.append(idx)
                    wb = load_workbook(output_filename)
                    ws = wb['Keyword Matches']
                    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                    for idx in new_row_indices:
                        excel_row = idx + 2
                        for cell in ws[excel_row]:
                            cell.fill = yellow_fill
                    wb.save(output_filename)
                    logging.info(f"Highlighted {len(new_row_indices)} new rows in yellow based on PostedDate newer than previous file timestamp: {prev_timestamp_str}")
                    if parse_failures > 0:
                        logging.info(f"Could not parse PostedDate for {parse_failures} rows. Examples: {failed_examples}")
                else:
                    logging.info(f"Could not extract timestamp from previous Excel filename for highlighting: {prev_file}")
            else:
                logging.info("No previous Excel file found for row comparison/highlighting.")
        except Exception as e:
            logging.info(f"Failed to highlight new rows: {e}")
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