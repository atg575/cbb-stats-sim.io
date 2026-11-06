import os
import requests
from bs4 import BeautifulSoup
import pandas as pd
from io import StringIO
from openpyxl import load_workbook

# Create a session with enhanced headers
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
    "Referer": "https://kenpom.com",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive"
})

session.cookies.set("PHPSESSID", "2af90006f87506a3380d62647c82fcdc")

def convert_numeric_columns(df):
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except (ValueError, TypeError):
            pass
    return df

def clean_string_columns(df):
    """Standardizes string columns to prevent em/en-dashes from causing encoding issues."""
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(
            lambda x: str(x).replace('\xa0', ' ').replace('—', '-').replace('–', '-').strip() if pd.notnull(x) else x
        )
    return df

def scrape_and_save_to_excel(file_name="2026-27 Men's College Basketball Stats_code.xlsx"):
    url = "https://kenpom.com/index.php"
    flstudio = [1]

    for a in flstudio:
        try:
            response = session.get(url, timeout=10)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data: {e}")
            return

        # Use .content to handle UTF-8 properly just in case
        soup = BeautifulSoup(response.content, 'html.parser')

        table = soup.find('table', {'id': 'ratings-table'})
        
        if not table:
            print("Table not found!")
            return

        try:
            df = pd.read_html(StringIO(str(table)))[0]
        except ValueError:
            print("Error parsing table. It might be empty or malformed.")
            return

        # Flatten MultiIndex columns if needed
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = ['_'.join(col).strip() for col in df.columns.values]

        # Process clean data types
        df = clean_string_columns(df)
        df = convert_numeric_columns(df)

        sheet_name = str("Kenpom Ratings")

        if os.path.exists(file_name):
            with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Data replaced in {file_name}")
        else:
            with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Data saved to new file {file_name}")

if __name__ == "__main__":
    scrape_and_save_to_excel()