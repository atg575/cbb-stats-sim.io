import os
import time
import pandas as pd
from io import StringIO
from datetime import datetime
from playwright.sync_api import sync_playwright

def scrape_espn_bpi(out_path="data/raw/espn_bpi.csv"):
    base_url = "https://www.espn.com/mens-college-basketball/bpi/_/page/"
    all_dfs = []
    
    print("Launching browser to bypass ESPN bot protection...")
    
    with sync_playwright() as p:
        # Launch browser (Headless=True means it runs invisibly in the background)
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        # ESPN BPI has roughly 362 teams. At 25 teams per page, that's up to 15 pages.
        for page_num in range(1, 16):
            url = f"{base_url}{page_num}"
            print(f"Fetching ESPN BPI Page {page_num}/15...")
            
            try:
                page.goto(url, wait_until='domcontentloaded')
                
                # Wait for the table to appear (up to 15 seconds)
                page.wait_for_selector("table", timeout=5000)
                
                # Extract the sanitized HTML directly from the browser
                html_content = page.content()
                
                # Because the browser sanitizes the HTML, Pandas will no longer throw an XMLSyntaxError
                tables = pd.read_html(StringIO(html_content))
                
                if len(tables) >= 2:
                    # ESPN visually splits its tables into a left side (Names) and right side (Stats)
                    left_table = tables[0]
                    right_table = tables[1]
                    merged = pd.concat([left_table, right_table], axis=1)
                    all_dfs.append(merged)
                elif len(tables) == 1:
                    all_dfs.append(tables[0])
                else:
                    print(f"No tables found on page {page_num}. Ending scrape.")
                    break
                    
            except Exception as e:
                # If it times out, we likely hit the end of the pages (e.g., page 15 might be empty)
                print(f"Finished scraping or timed out on page {page_num}: {e}")
                break
                
            time.sleep(1) # Brief pause between page loads
            
        browser.close()
        
    if not all_dfs:
        print("Failed to scrape any data.")
        return
        
    # Combine all pages into one massive DataFrame
    final_df = pd.concat(all_dfs, ignore_index=True)
        
    # Flatten Multi-level column headers if ESPN uses them
    if isinstance(final_df.columns, pd.MultiIndex):
        final_df.columns = ['_'.join(col).strip() for col in final_df.columns.values]
    else:
        final_df.columns = [str(col).strip() for col in final_df.columns]
        
    # Clean up mid-table headers ESPN might have injected
    team_col = next((col for col in final_df.columns if 'team' in str(col).lower()), None)
    if team_col:
         final_df = final_df[final_df[team_col].astype(str).str.lower() != 'team']
    
    # Clean string columns (replace em/en-dashes and non-breaking spaces)
    for col in final_df.select_dtypes(include=['object']).columns:
        final_df[col] = final_df[col].apply(
            lambda x: str(x).replace('\xa0', ' ').replace('—', '-').replace('–', '-').strip() if pd.notnull(x) else x
        )
    
    # Force Numeric conversion so Excel formulas don't break
    ignore_cols = ['team', 'conf', 'school']
    for col in final_df.columns:
        is_text_col = any(text_col in str(col).lower() for text_col in ignore_cols)
        if not is_text_col:
            final_df[col] = pd.to_numeric(final_df[col], errors='coerce')
            final_df[col] = final_df[col].fillna(0)
            
    # Add extraction date
    final_df['Date_Pulled'] = datetime.now().strftime('%Y-%m-%d')
            
    print(f"\nTotal teams extracted: {len(final_df)}")
    
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    final_df.to_csv(out_path, index=False)
    print(f"Saved {len(final_df)} rows to {out_path}")

if __name__ == '__main__':
    scrape_espn_bpi()