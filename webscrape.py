import os
import requests

def main():
    print("SAM.gov Webscrape")
    print("=================")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.join(script_dir, "input")
    output_folder = os.path.join(script_dir, "output")
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    # Read URL from sam_url.txt
    url_file = os.path.join(script_dir, "sam_url.txt")
    try:
        with open(url_file, "r") as uf:
            url = uf.readline().strip()
    except Exception as e:
        print(f"Failed to read URL from {url_file}: {e}")
        return

    csv_filename = os.path.join(input_folder, "ContractOpportunitiesFullCSV.csv")
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

if __name__ == "__main__":
    main()
