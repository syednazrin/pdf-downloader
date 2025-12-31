import pandas as pd
import requests
import os
import re
import urllib3

# === SETTINGS ===
excel_file = "URL links.xlsx"      # Excel file name (in same folder)
sheet_name = "Latest"                # 0 = first sheet
start_row = 4                 # e.g., 1 for C1
end_row = 93                  # e.g., 10 for C10
column_letter = "F"           # Excel column with URLs
download_folder = "downloads" # Folder to save PDFs

# === SETUP ===
os.makedirs(download_folder, exist_ok=True)

# Disable SSL warnings (for sites with certificate issues)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Calculate number of rows to read
rows_to_read = end_row - start_row + 1

# Read only the specified range from Excel
df = pd.read_excel(
    excel_file,
    sheet_name=sheet_name,
    usecols=column_letter,
    skiprows=range(start_row - 1),  # skip rows before start_row
    nrows=rows_to_read
)

# === DOWNLOAD LOOP ===
for i, url in enumerate(df.iloc[:, 0], start=start_row):
    if pd.isna(url):
        continue

    # Clean filename
    filename = f"file_{i}.pdf"
    filepath = os.path.join(download_folder, filename)

    try:
        print(f"⬇️ Downloading {url} ...")

        # Enhanced headers to bypass 403 Forbidden
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/115.0 Safari/537.36"
            ),
            "Referer": "https://www.manulifeim.com.my/",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Connection": "keep-alive"
        }

        # Download PDF
        response = requests.get(url, headers=headers, timeout=20, verify=False)
        response.raise_for_status()

        # Save PDF
        with open(filepath, "wb") as f:
            f.write(response.content)

        print(f"✅ Saved: {filepath}")

    except Exception as e:
        print(f"❌ Error downloading {url}: {e}")

print("\n🎉 All downloads completed!")
