import os
import requests
import pandas as pd
from io import StringIO

# Create a session with enhanced headers
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
    "Referer": "https://kenpom.com",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive"
})

# Your active KenPom login cookie is required to access the download link
session.cookies.set("PHPSESSID", os.environ.get("KENPOM_PHPSESSID", "ae7f710eff40f888146979b5b3c56989"))

def clean_string_columns(df):
    """Standardizes string columns to prevent em/en-dashes from causing encoding issues."""
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(
            lambda x: str(x).replace('\xa0', ' ').replace('—', '-').replace('–', '-').strip() if pd.notnull(x) else x
        )
    return df

def force_numeric_columns(df):
    """Forces all stat and rank columns into true numbers while keeping Team, Conf, etc as text."""
    text_cols = ['team', 'conf', 'w-l', 'record']
    for col in df.columns:
        if not any(t in str(col).lower() for t in text_cols):
            cleaned = df[col].astype(str).str.replace('+', '', regex=False).str.replace(',', '', regex=False).str.strip()
            df[col] = pd.to_numeric(cleaned, errors='coerce').fillna(0)
    return df

def scrape_and_save_to_csv(out_path="data/raw/height.csv"):
    # Target the direct CSV download link
    url = "https://kenpom.com/getdata.php?file=height27"

    try:
        response = session.get(url, timeout=10)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        return

    # Check if the response is actually the CSV data or a redirect to the login page
    if "Subscribe Now" in response.text or "<html" in response.text.lower():
         print("Error: Received a webpage instead of CSV data. Your PHPSESSID cookie may be expired.")
         return

    try:
        # Read the raw CSV text directly into Pandas
        df = pd.read_csv(StringIO(response.text))
    except ValueError:
        print("Error parsing CSV data. It might be empty or malformed.")
        return

    # Clean and format the data
    df = clean_string_columns(df)
    df = force_numeric_columns(df)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    df.to_csv(out_path, index=False)
    print(f"Saved {len(df)} rows to {out_path}")

if __name__ == "__main__":
    scrape_and_save_to_csv()