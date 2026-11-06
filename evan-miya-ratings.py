import pandas as pd
import time
import os
import re
from playwright.sync_api import sync_playwright

def clean_dataframe(df):
    """Removes injury artifacts, converts em/en dashes, and strips trailing spaces."""
    if df.empty or 'Team' not in df.columns:
        return df
    
    # 1. Remove the specific newline + bandage emoji
    df['Team'] = df['Team'].astype(str).str.replace(r'\n🤕', '', regex=True)
    
    # 2. Standardize dashes to prevent encoding glitches
    df['Team'] = df['Team'].str.replace('—', '-').str.replace('–', '-')
    
    # 3. Remove any other lingering non-ASCII characters (emojis) just to be safe
    df['Team'] = df['Team'].astype(str).str.replace(r'[^\x00-\x7F]+', '', regex=True)
    
    # 4. Strip whitespace
    df['Team'] = df['Team'].str.strip()
    return df

def extract_grid_data(page):
    script = """
    () => {
        let data = [];
        let rows = Array.from(document.querySelectorAll('[role="row"], .rt-tr, tr'));
        
        if (rows.length === 0) return []; 

        rows.forEach(row => {
            let cols = Array.from(row.querySelectorAll('[role="columnheader"], [role="cell"], .rt-td, .rt-th, td, th'));
            if (cols.length > 0) {
                data.push(cols.map(col => col.innerText.trim()));
            }
        });
        return data;
    }
    """
    raw_data = page.evaluate(script)
    if not raw_data:
        return pd.DataFrame()
        
    headers = raw_data[0]
    rows = raw_data[1:]
    return pd.DataFrame(rows, columns=headers)

def get_evanmiya_data_persistent():
    with sync_playwright() as p:
        user_data_dir = os.path.join(os.getcwd(), "playwright_profile")
        
        browser = p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=True 
        )
        
        page = browser.pages[0]
        
        print("Loading EvanMiya Team Ratings...")
        page.goto("https://evanmiya.com/team-ratings")
        
        print("Waiting for data to render...")
        page.wait_for_selector("text=Rank", timeout=15000)
        
        print("Expanding table to show 1000 rows...")
        try:
            page.locator("select").filter(has_text="1000").select_option("1000")
            print("Waiting for 1000 rows to load...")
            time.sleep(3)
        except Exception as e:
            print("Could not auto-select 1000 rows. It may already be expanded or requires manual clicking.")
        
        print("Extracting Regular Ratings...")
        regular_df = extract_grid_data(page)
        regular_df = clean_dataframe(regular_df)
        
        print("Switching to Injured Ratings...")
        try:
            dropdown_container = page.get_by_text("Include Current Injuries?").locator("xpath=..")
            dropdown_container.get_by_text("No").click()
            time.sleep(1) 
            page.get_by_text("Yes", exact=True).click()
            
            print("Waiting for injured data to load...")
            time.sleep(5) 
            
            injured_df = extract_grid_data(page)
            injured_df = clean_dataframe(injured_df)
        except Exception as e:
            print(f"Could not switch to injured ratings: {e}")
            injured_df = pd.DataFrame()
            
        browser.close()
        return regular_df, injured_df

if __name__ == "__main__":
    reg_df, inj_df = get_evanmiya_data_persistent()
    file_name="2026-27 Men's College Basketball Stats_code.xlsx"
    if reg_df.empty:
        print("Failed to extract data.")
    else:
        print(f"Saving data to {file_name}...")
        reg_sheet_name = str("EvanMiya Regular Ratings")
        inj_sheet_name = str("EvanMiya Injured Ratings")

        if os.path.exists(file_name):
            with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                reg_df.to_excel(writer, sheet_name=reg_sheet_name, index=False)
                inj_df.to_excel(writer, sheet_name=inj_sheet_name, index=False)
                print(f"Data replaced in {file_name}")
        else:
            with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                reg_df.to_excel(writer, sheet_name=reg_sheet_name, index=False)
                inj_df.to_excel(writer, sheet_name=inj_sheet_name, index=False)
                print(f"Data saved to new file {file_name}")
      
        print(f"Extraction complete! Pulled {len(reg_df)} regular rows and {len(inj_df)} injured rows.")