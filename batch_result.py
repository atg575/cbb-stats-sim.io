import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

# Your exact Excel logic translated into Python
def get_accuracy(pred_away, pred_home, act_away, act_home):
    pred_margin = pred_away - pred_home
    act_margin = act_away - act_home
    
    pred_total = pred_away + pred_home
    act_total = act_away + act_home
    
    pred_winner = "Away" if pred_away > pred_home else "Home" if pred_home > pred_away else "Tie"
    act_winner = "Away" if act_away > act_home else "Home" if act_home > act_away else "Tie"
    
    margin_diff = abs(pred_margin - act_margin)
    max_margin = max(abs(pred_margin), abs(act_margin))
    margin_pct = (margin_diff / max_margin) if max_margin != 0 else 0
    
    # 5 point margin difference OR 40% margin ratio threshold
    margin_accurate = (margin_diff <= 5) or (margin_pct <= 0.4)
    # 10 point total score threshold
    total_accurate = abs(pred_total - act_total) <= 10
    winner_accurate = (pred_winner == act_winner)
    
    if margin_accurate and total_accurate and winner_accurate: return "ALL"
    elif margin_accurate and winner_accurate: return "TEAM+DIFF"
    elif winner_accurate: return "TEAM"
    elif total_accurate: return "TOTAL"
    else: return "WRONG"

# Your exact Certainty thresholds translated into Python
def get_certainty(prob_away, prob_home):
    max_prob = max(prob_away, prob_home)
    if max_prob >= 0.90: return "90%+"
    elif max_prob >= 0.75: return "75%+"
    elif max_prob >= 0.60: return "60%+"
    elif max_prob >= 0.50: return "50%+"
    else: return "-"

def grade_games():
    url_date = input("Enter date to grade (Ex. 2026-11-04): ").strip()
    file_name = "Batch_Simulation_Results.xlsx"
    
    if not os.path.exists(file_name):
        print(f"File '{file_name}' not found!")
        return

    try:
        df = pd.read_excel(file_name, sheet_name=url_date)
    except Exception as e:
        print(f"Could not find sheet '{url_date}' in '{file_name}'.")
        return

    print(f"\nFetching actual scores from Warren Nolan for {url_date}...")
    url = f"https://www.warrennolan.com/basketball/2027/scores-live?type=All%20Games&date={url_date}"
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(url, headers=headers)
    soup = BeautifulSoup(resp.content, 'html.parser')
    
    # Safely extract all text on the webpage to scan for the actual scores
    text_tokens = list(soup.stripped_strings)
    clean_tokens = [t.lower().strip() for t in text_tokens]
    
    scores_dict = {}
    teams_in_spreadsheet = df['School'].dropna().astype(str).tolist()
    
    # Locate the teams from your spreadsheet on the webpage and grab their actual score
    for team in teams_in_spreadsheet:
        team_lower = team.lower().strip()
        if team_lower in clean_tokens:
            idx = clean_tokens.index(team_lower)
            # Find the numerical score sitting right next to the team name in the HTML
            for i in range(idx + 1, min(idx + 5, len(clean_tokens))):
                if clean_tokens[i].isdigit():
                    scores_dict[team_lower] = int(clean_tokens[i])
                    break

    # Add new column for Actual Scores
    if "Actual Score" not in df.columns:
        df["Actual Score"] = ""
    
    # Pair up your model names with the columns they use
    models = [
        ("Live Morey Win %", "Live Morey Score", "Live Morey"),
        ("Consensus Morey Win %", "Consensus Morey Score", "Consensus Morey"),
        ("Live Log5 Win %", "Live Log5 Score", "Live Log5"),
        ("Consensus Log5 Win %", "Consensus Log5 Score", "Consensus Log5"),
        ("Adj Eff Win %", "Adj Eff Score", "Adj Eff"),
        ("Regular Win %", "Regular Score", "Regular")
    ]
    
    # Prepare the accuracy and certainty columns for all 6 models
    for _, _, model_name in models:
        if f"{model_name} Accuracy" not in df.columns:
            df[f"{model_name} Accuracy"] = ""
            df[f"{model_name} Certainty"] = ""
        
    # Iterate through the games (steps by 3 to account for Away, Home, and the Empty row)
    for i in range(0, len(df), 3):
        if i + 1 >= len(df):
            break
            
        away_team = str(df.at[i, "School"]).strip().lower()
        home_team = str(df.at[i+1, "School"]).strip().lower()
        
        if pd.isna(df.at[i, "School"]) or away_team == 'nan':
            continue

        act_away = scores_dict.get(away_team)
        act_home = scores_dict.get(home_team)

        if act_away is None or act_home is None:
            print(f"Skipping {away_team.title()} vs {home_team.title()} - Final scores not found/game still active.")
            continue

        df.at[i, "Actual Score"] = act_away
        df.at[i+1, "Actual Score"] = act_home
        
        # Calculate grades for every single model and append to the spreadsheet
        for win_col, score_col, model_name in models:
            try:
                pred_away_score = float(df.at[i, score_col])
                pred_home_score = float(df.at[i+1, score_col])
                pred_away_win = float(df.at[i, win_col])
                pred_home_win = float(df.at[i+1, win_col])
                
                accuracy = get_accuracy(pred_away_score, pred_home_score, act_away, act_home)
                certainty = get_certainty(pred_away_win, pred_home_win)
                
                df.at[i, f"{model_name} Accuracy"] = accuracy
                df.at[i, f"{model_name} Certainty"] = certainty
            except (ValueError, TypeError):
                continue 

    # Clean up empty columns Pandas renamed to 'Unnamed: X' to preserve your visual spacing
    new_columns = []
    for col in df.columns:
        if str(col).startswith("Unnamed:"):
            new_columns.append("")
        else:
            new_columns.append(col)
    df.columns = new_columns

    # Save the graded data back over the same daily sheet
    with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=url_date, index=False)
        
    print(f"\nSuccessfully graded '{url_date}' and updated {file_name}!")

if __name__ == "__main__":
    grade_games()