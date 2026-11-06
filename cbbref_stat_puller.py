from datetime import datetime
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
    "Referer": "https://www.sports-reference.com",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive"
})

def convert_numeric_columns(df):
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='ignore')
    return df

def clean_string_columns(df):
    # Standardizes all string columns to prevent trailing spaces, weird unicode spaces, and em/en-dashes
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(
            lambda x: str(x).replace('\xa0', ' ').replace('—', '-').replace('–', '-').strip() if pd.notnull(x) else x
        )
    return df

def scrape_and_save_to_excel(file_name="2026-27 Men's College Basketball Stats_code.xlsx"):
    url_a = "https://www.sports-reference.com/cbb/seasons/men/2027-school-stats.html#basic_school_stats::1"
    url_b = "https://www.sports-reference.com/cbb/seasons/men/2027-opponent-stats.html#basic_opp_stats::1"

    try:
        response_a = session.get(url_a, timeout=10)
        response_b = session.get(url_b, timeout=10)
        response_a.raise_for_status()
        response_b.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        return

    soup_a = BeautifulSoup(response_a.content, 'html.parser')
    soup_b = BeautifulSoup(response_b.content, 'html.parser')

    table_a = soup_a.find('table', {'id': 'basic_school_stats'})
    table_b = soup_b.find('table', {'id': 'basic_opp_stats'})

    if not table_a or not table_b:
        print("One or both tables not found!")
        return

    try:
        df_a = pd.read_html(StringIO(str(table_a)))[0]
        df_b = pd.read_html(StringIO(str(table_b)))[0]
    except ValueError:
        print("Error parsing tables. They might be empty or malformed.")
        return

    df_a['Date'] = datetime.now().strftime('%Y-%m-%d')
    df_b['Date'] = datetime.now().strftime('%Y-%m-%d')

    if isinstance(df_a.columns, pd.MultiIndex):
        df_a.columns = ['_'.join(col).strip() for col in df_a.columns.values]
    if isinstance(df_b.columns, pd.MultiIndex):
        df_b.columns = ['_'.join(col).strip() for col in df_b.columns.values]

    df_a = convert_numeric_columns(df_a)
    df_b = convert_numeric_columns(df_b)
    
    df_a = clean_string_columns(df_a)
    df_b = clean_string_columns(df_b)

    sheet_name_a = "School Stats"
    sheet_name_b = "Opponent Stats"

    if os.path.exists(file_name):
        with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df_a.to_excel(writer, sheet_name=sheet_name_a, index=False)
            df_b.to_excel(writer, sheet_name=sheet_name_b, index=False)
            print(f"Data replaced in {file_name}")
    else:
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_a.to_excel(writer, sheet_name=sheet_name_a, index=False)
            df_b.to_excel(writer, sheet_name=sheet_name_b, index=False)
            print(f"Data saved to new file {file_name}")

if __name__ == "__main__":
    scrape_and_save_to_excel()  # Optionally run once immediately
