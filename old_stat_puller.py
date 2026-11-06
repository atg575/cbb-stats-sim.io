from datetime import datetime
import os
import requests
from bs4 import BeautifulSoup
import pandas as pd
from io import StringIO

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Referer": "https://www.sports-reference.com",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive"
})

def convert_numeric_columns(df):
    """Forces all stat columns into actual numeric types so Excel formats them as numbers."""
    ignore_cols = ['school', 'conf', 'date', 'match_school']
    for col in df.columns:
        # Check if the column name contains any of the known text headers
        is_text_col = any(text_col in str(col).lower() for text_col in ignore_cols)
        
        if not is_text_col:
            # Force to numeric (turns invalid text like 'Totals' or '-' into NaN)
            df[col] = pd.to_numeric(df[col], errors='coerce')
            # Fill NaNs with 0 to ensure clean mathematical operations
            df[col] = df[col].fillna(0)
            
    return df

def clean_string_columns(df):
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(
            lambda x: str(x).replace('\xa0', ' ').replace('—', '-').replace('–', '-').strip() if pd.notnull(x) else x
        )
    return df

def remove_mid_table_headers(df):
    school_col = next((col for col in df.columns if 'School' in col), None)
    if school_col:
        df = df[df[school_col].astype(str).str.strip().str.lower() != 'school']
    return df.reset_index(drop=True)

def ensure_all_teams_exist(df, hist_df):
    if hist_df is None or hist_df.empty:
        return df
        
    school_col = next((col for col in df.columns if 'School' in col), 'School')
    
    if not df.empty and school_col in df.columns:
        current_teams = set(df[school_col].astype(str).str.replace(' NCAA', '').str.strip().str.lower())
    else:
        current_teams = set()
        
    missing_teams = []
    for school in hist_df['School'].unique():
        clean_hist_school = str(school).replace(' NCAA', '').strip().lower()
        if clean_hist_school not in current_teams:
            new_row = {col: 0 for col in df.columns} 
            new_row[school_col] = school
            missing_teams.append(new_row)
            
    if missing_teams:
        missing_df = pd.DataFrame(missing_teams)
        df = pd.concat([df, missing_df], ignore_index=True)
        
    return df

def safe_float(val):
    if pd.isna(val) or val == '':
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def apply_blended_totals(df, hist_df, is_opp=False):
    """
    Pads current season totals with historical averages for teams with < 5 games.
    If a team has no history (like a new D1 team), pads with League Averages.
    Outputs the final blended TOTAL so the user's spreadsheet can divide by Games.
    """
    if df is None or df.empty:
        return df

    # --- 1. Calculate League Averages for fallback padding ---
    league_averages = {}
    stat_categories = ['FG', 'FGA', '3P', '3PA', 'FT', 'FTA', 'ORB', 'TRB', 'AST', 'STL', 'BLK', 'TOV', 'PF']
    pts_mappings = [
        ('Tm.', 'Opp.' if is_opp else 'Tm.'), 
        ('Opp.', 'Tm.' if is_opp else 'Opp.'),
        ('PTS', 'Opp.' if is_opp else 'Tm.')
    ]
    
    g_col = next((col for col in df.columns if col.endswith('_G') or col == 'G'), None)
    if g_col:
        for stat in stat_categories + [m[0] for m in pts_mappings]:
            curr_col = next((c for c in df.columns if c.endswith(f'_{stat}') or c == stat), None)
            if curr_col:
                total_stat = pd.to_numeric(df[curr_col], errors='coerce').sum()
                total_games = pd.to_numeric(df[g_col], errors='coerce').sum()
                league_averages[stat] = (total_stat / total_games) if total_games > 0 else 0.0

    # --- 2. Clean Historical Names ---
    if hist_df is not None and not hist_df.empty:
        hist_df['Match_School'] = hist_df['School'].astype(str)\
            .str.replace('\xa0', ' ')\
            .str.replace(' NCAA', '')\
            .str.replace(' NIT', '')\
            .str.replace(' CBI', '')\
            .str.replace('—', '-')\
            .str.replace('–', '-')\
            .str.strip().str.lower()
        
        # Inject any historical teams that haven't played yet this year
        df = ensure_all_teams_exist(df, hist_df)

    if not g_col:
        return df

    # --- 3. Clean Current Names ---
    school_col = next((col for col in df.columns if 'School' in col), None)
    df['Match_School'] = df[school_col].astype(str)\
        .str.replace('\xa0', ' ')\
        .str.replace(' NCAA', '')\
        .str.replace(' NIT', '')\
        .str.replace(' CBI', '')\
        .str.replace('—', '-')\
        .str.replace('–', '-')\
        .str.strip().str.lower()
    
    # --- 4. Process EVERY team (including new D1 teams like West Florida) ---
    for idx, row in df.iterrows():
        g = safe_float(row[g_col])
        school_match = row['Match_School']
        has_hist = False
        
        # Check if they have historical data
        if hist_df is not None and not hist_df.empty:
            hist_row = hist_df[hist_df['Match_School'] == school_match]
            if not hist_row.empty:
                has_hist = True
                hist_row = hist_row.iloc[0]

        for stat in stat_categories + [m[0] for m in pts_mappings]:
            curr_col = next((c for c in df.columns if c.endswith(f'_{stat}') or c == stat), None)
            
            if stat in [m[0] for m in pts_mappings]:
                hist_stat_name = next(m[1] for m in pts_mappings if m[0] == stat)
            else:
                hist_stat_name = f"Opp_{stat}" if is_opp else stat
            
            if curr_col:
                curr_total = safe_float(row[curr_col])

                if g < 5:
                    missing_g = 5 - g
                    
                    # USE HISTORY IF IT EXISTS. OTHERWISE USE LEAGUE AVERAGE.
                    if has_hist and hist_stat_name in hist_row and pd.notna(hist_row[hist_stat_name]):
                        hist_per_game = safe_float(hist_row[hist_stat_name]) / 33.0
                    else:
                        hist_per_game = league_averages.get(stat, 0.0)
                    
                    blended_total = curr_total + (missing_g * hist_per_game)
                    df.at[idx, curr_col] = round(blended_total, 1)
                else:
                    df.at[idx, curr_col] = round(curr_total, 1)

        pct_mappings = [('FG', 'FGA', 'FG%'), ('3P', '3PA', '3P%'), ('FT', 'FTA', 'FT%')]
        for makes, attempts, pct in pct_mappings:
            m_col = next((c for c in df.columns if c.endswith(f'_{makes}') or c == makes), None)
            a_col = next((c for c in df.columns if c.endswith(f'_{attempts}') or c == attempts), None)
            p_col = next((c for c in df.columns if c.endswith(f'_{pct}') or c == pct), None)
            
            if m_col and a_col and p_col:
                a_val = safe_float(df.at[idx, a_col])
                m_val = safe_float(df.at[idx, m_col])
                if a_val > 0:
                    df.at[idx, p_col] = round(m_val / a_val, 3)

        if g < 5:
            df.at[idx, g_col] = 5

    return df.drop(columns=['Match_School'])

def scrape_and_save_to_excel(file_name="2026-27 Men's College Basketball Stats_code.xlsx"):
    url_a = "https://www.sports-reference.com/cbb/seasons/men/2027-school-stats.html#basic_school_stats::1"
    url_b = "https://www.sports-reference.com/cbb/seasons/men/2027-opponent-stats.html#basic_opp_stats::1"

    hist_file = "CBB_3_Year_Team_And_Opponent_Averages.csv"
    hist_df = None
    if os.path.exists(hist_file):
        hist_df = pd.read_csv(hist_file)
    else:
        print(f"Warning: '{hist_file}' not found. Teams cannot be padded.")

    try:
        response_a = session.get(url_a, timeout=10)
        response_b = session.get(url_b, timeout=10)
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        return

    soup_a = BeautifulSoup(response_a.content, 'html.parser')
    soup_b = BeautifulSoup(response_b.content, 'html.parser')

    table_a = soup_a.find('table', {'id': 'basic_school_stats'})
    table_b = soup_b.find('table', {'id': 'basic_opp_stats'})

    try:
        df_a = pd.read_html(StringIO(str(table_a)))[0]
        df_b = pd.read_html(StringIO(str(table_b)))[0]
    except ValueError:
        return

    if isinstance(df_a.columns, pd.MultiIndex):
        df_a.columns = ['_'.join(col).strip() for col in df_a.columns.values]
    if isinstance(df_b.columns, pd.MultiIndex):
        df_b.columns = ['_'.join(col).strip() for col in df_b.columns.values]

    df_a = remove_mid_table_headers(df_a)
    df_b = remove_mid_table_headers(df_b)

    df_a['Date'] = datetime.now().strftime('%Y-%m-%d')
    df_b['Date'] = datetime.now().strftime('%Y-%m-%d')

    df_a = clean_string_columns(df_a)
    df_b = clean_string_columns(df_b)
    
    # Run the strict numeric conversion before mathematically padding the totals
    df_a = convert_numeric_columns(df_a)
    df_b = convert_numeric_columns(df_b)
    
    df_a = apply_blended_totals(df_a, hist_df, is_opp=False)
    df_b = apply_blended_totals(df_b, hist_df, is_opp=True)

    if os.path.exists(file_name):
        with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df_a.to_excel(writer, sheet_name="School Stats", index=False)
            df_b.to_excel(writer, sheet_name="Opponent Stats", index=False)
    else:
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_a.to_excel(writer, sheet_name="School Stats", index=False)
            df_b.to_excel(writer, sheet_name="Opponent Stats", index=False)
            
    print("Success: Stats saved as true Numbers and Blended Totals.")

if __name__ == "__main__":
    scrape_and_save_to_excel()