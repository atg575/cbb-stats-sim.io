import requests
import pandas as pd
from io import StringIO
import ssl
from datetime import datetime
import os

# Ignore SSL verification if running into macOS certificate issues
ssl._create_default_https_context = ssl._create_unverified_context

def clean_string_columns(df):
    """Standardizes string columns to prevent em/en-dashes from causing encoding issues."""
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(
            lambda x: str(x).replace('\xa0', ' ').replace('—', '-').replace('–', '-').strip() if pd.notnull(x) else x
        )
    return df

def force_numeric_columns(df):
    """Forces all non-team/date columns into strict numeric types to prevent text formatting in Excel."""
    ignore_cols = ['team', 'school']
    for col in df.columns:
        is_text_col = any(text_col in str(col).lower() for text_col in ignore_cols)
        if not is_text_col:
            # Coerce forces invalid text into NaN, which we then fill with 0
            df[col] = pd.to_numeric(df[col], errors='coerce')
            df[col] = df[col].fillna(0)
    return df

def scrape_warren_nolan_elo(file_name="2026-27 Men's College Basketball Stats_code.xlsx"):
    url = "https://www.warrennolan.com/basketball/2027/elo"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.warrennolan.com/",
    }

    print(f"Fetching Elo Ratings from: {url}")
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        return

    try:
        # Use pandas to extract the tables directly from the HTML
        tables = pd.read_html(StringIO(response.text))
        
        if not tables:
            print("No tables found on the page.")
            return
            
        # The main data is usually in the first table
        df = tables[0]
        
        # Clean up column names (remove multi-indexing if present or strip spaces)
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = ['_'.join(col).strip() for col in df.columns.values]
        else:
            df.columns = [str(col).strip() for col in df.columns]
            
        # Clean team names and force numeric conversion to avoid text artifacts
        df = clean_string_columns(df)
        df = force_numeric_columns(df)

        # Add a date column for tracking when you pulled the data
        df['Date_Pulled'] = datetime.now().strftime('%Y-%m-%d')
        
        sheet_name = "Warren ELO"
        
        # Save to Excel
        if os.path.exists(file_name):
            with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Data replaced in {file_name}")
        else:
            with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Data saved to new file {file_name}")
            
        print(f"Success! Extracted {len(df)} teams and saved to '{file_name}'.")

    except ValueError:
        print("Error parsing the HTML table. The webpage format may have changed or the table is empty.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    scrape_warren_nolan_elo()