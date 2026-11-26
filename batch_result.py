import json
import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import sys

# Allows non-interactive cloud runs while keeping local prompt support
if len(sys.argv) > 1:
    url_date = sys.argv[1]
elif os.environ.get("GITHUB_ACTIONS") == "true":
    url_date = (datetime.now(ZoneInfo("America/New_York")) - timedelta(hours=4)).strftime("%Y-%m-%d")
else:
    if input("PAST DATE GRADING MODE? (y/n): ").strip().lower() == 'y':
        url_date = input("Enter the date to grade (YYYY-MM-DD): ").strip()
    else:
        url_date = (datetime.now(ZoneInfo("America/New_York")) - timedelta(hours=4)).strftime("%Y-%m-%d")


def normalize_name(name):
    """Normalizes team names so minor punctuation/dash differences never skip a game."""
    if pd.isna(name):
        return ""
    s = str(name).lower().strip()
    s = s.replace('\xa0', ' ').replace('—', '-').replace('–', '-')
    s = re.sub(r"[^\w\s-]", "", s)
    return re.sub(r"\s+", " ", s).strip()


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
    
    margin_exact = (margin_diff == 0)
    total_exact = (pred_total == act_total)
    margin_accurate = (margin_diff <= 5) or (margin_pct <= 0.4)
    total_accurate = abs(pred_total - act_total) <= 10
    winner_accurate = (pred_winner == act_winner)
    
    if margin_exact and total_exact and winner_accurate: return "SUPER ALL"
    elif margin_accurate and total_accurate and winner_accurate: return "ALL"
    elif margin_accurate and winner_accurate: return "TEAM AND DIFF"
    elif winner_accurate: return "TEAM"
    elif total_accurate: return "TOTAL"
    else: return "WRONG"


def get_certainty(prob_away, prob_home):
    max_prob = max(prob_away, prob_home)
    if max_prob >= 0.90: return "90%+"
    elif max_prob >= 0.75: return "75%+"
    elif max_prob >= 0.60: return "60%+"
    elif max_prob >= 0.50: return "50%+"
    else: return "-"


def is_strictly_final(tokens):
    """
    Returns True ONLY if the game box tokens explicitly contain a 'FINAL' status
    and do NOT contain any live game clock or in-progress period indicators.
    """
    low_tokens = [t.lower().strip() for t in tokens if t.strip()]
    
    # 1. Must contain an explicit 'final' status token (e.g. 'final', 'final/ot', 'final - 2ot')
    has_final_token = any(
        t == "final" or t.startswith("final/") or t.startswith("final ") or t.startswith("final-")
        for t in low_tokens
    )
    if not has_final_token:
        return False

    # 2. Reject immediately if any token looks like a running game clock (e.g., '14:22', '0:45.2')
    # or a scheduled start time ('7:00 pm', '1st half', '2nd half', 'halftime')
    live_keywords = {"halftime", "half", "1st half", "2nd half", "in progress", "live", "delayed", "postponed", "canceled"}
    for t in low_tokens:
        if t in live_keywords:
            return False
        # Matches running game clocks like '12:34' or '4:05'
        if re.match(r"^\d{1,2}:\d{2}(\.\d+)?$", t):
            return False

    return True

def extract_verified_total_from_numbers(nums):
    """
    Given a list of integers from a completed team row (which might start with a Top-25 rank
    followed by [1st Half, 2nd Half, (OT...), Total]), verifies that each half is realistic
    and that the periods sum to the Total score.
    """
    if len(nums) < 3:
        return None

    for end_idx in range(len(nums), 2, -1):
        candidate_total = nums[end_idx - 1]
        if candidate_total < 10:
            continue
        for start_idx in range(0, end_idx - 2):
            periods = nums[start_idx:end_idx - 1]
            # A completed game must have at least 2 halves (first two periods >= 5 pts each)
            # and all periods must sum to candidate_total
            if len(periods) >= 2 and periods[0] >= 5 and periods[1] >= 5 and sum(periods) == candidate_total:
                return candidate_total

    return None


def scrape_completed_matchups(soup):
    """
    Extracts every individual game box from the page, verifies that the specific game box
    is strictly FINAL, and returns a list of completed game dictionaries:
    [{norm_team_1: score_1, norm_team_2: score_2}, ...]
    """
    completed_games = []

    # Inspect each individual <table> that does not contain sub-tables (each single game box)
    for table in soup.find_all("table"):
        if table.find("table"):
            continue

        # Include immediate header/status sibling or tight single-game wrapper (only if it has 1 table)
        game_container = table
        parent = table.find_parent("div")
        if parent and len(parent.find_all("table")) == 1:
            game_container = parent

        box_tokens = [t.strip() for t in game_container.stripped_strings if t.strip()]
        if not is_strictly_final(box_tokens):
            continue

        # Parse each team's <tr> inside this verified FINAL game table
        game_scores = {}
        for tr in table.find_all("tr"):
            row_tokens = [t.strip() for t in tr.stripped_strings if t.strip()]
            if not row_tokens:
                continue

            row_norm = [normalize_name(t) for t in row_tokens]
            # Find the integer sequence at the end of the row
            nums = [int(tok) for tok in row_tokens if tok.isdigit()]
            total_score = extract_verified_total_from_numbers(nums)

            if total_score is not None:
                # Record all non-numeric normalized strings in this row as potential team name keys
                for raw_tok, norm_tok in zip(row_tokens, row_norm):
                    if norm_tok and not raw_tok.isdigit() and not raw_tok.startswith("("):
                        game_scores[norm_tok] = total_score

        if len(game_scores) >= 2:
            completed_games.append(game_scores)

    return completed_games


def grade_games(url_date):
    sim_path = os.path.join("data", "sims", f"{url_date}.json")
    if not os.path.isfile(sim_path):
        print(f"Simulation file '{sim_path}' not found!")
        return

    with open(sim_path, "r", encoding="utf-8") as f:
        slate = json.load(f)

    games = slate.get("games", [])
    if not games:
        print(f"No games found in {sim_path}.")
        return

    print(f"\nFetching actual scores from Warren Nolan for {url_date}...")
    url = f"https://www.warrennolan.com/basketball/2027/scores-live?type=All%20Games&date={url_date}"
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(url, headers=headers, timeout=15)
    soup = BeautifulSoup(resp.content, 'html.parser')

    # Only scrape game boxes that are strictly FINAL
    completed_games = scrape_completed_matchups(soup)

    graded_count = 0
    for g in games:
        away_raw = str(g.get("t1", "")).strip()
        home_raw = str(g.get("t2", "")).strip()
        if not away_raw or not home_raw:
            continue

        norm_away = normalize_name(away_raw)
        norm_home = normalize_name(home_raw)

        act_away, act_home = None, None
        for game_box in completed_games:
            if norm_away in game_box and norm_home in game_box:
                act_away = game_box[norm_away]
                act_home = game_box[norm_home]
                break

        if act_away is None or act_home is None:
            g["s1"] = None
            g["s2"] = None
            for m_data in g.get("models", {}).values():
                m_data.pop("acc", None)
                m_data.pop("cert", None)
            print(f"Skipping {away_raw.title()} vs {home_raw.title()} - Game not Final yet.")
            continue

        g["s1"] = int(act_away)
        g["s2"] = int(act_home)
        graded_count += 1

        for m_name, m_data in g.get("models", {}).items():
            try:
                pred_away_score = float(m_data["ps1"])
                pred_home_score = float(m_data["ps2"])
                pred_away_win = float(m_data["wp1"])
                pred_home_win = float(m_data["wp2"])
                if pred_away_win > 1.0:
                    pred_away_win /= 100.0
                if pred_home_win > 1.0:
                    pred_home_win /= 100.0

                m_data["acc"] = get_accuracy(pred_away_score, pred_home_score, act_away, act_home)
                m_data["cert"] = get_certainty(pred_away_win, pred_home_win)
            except (KeyError, ValueError, TypeError):
                continue

    tmp_path = sim_path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump({"date": url_date, "games": games}, f, indent=2, ensure_ascii=False)
    os.replace(tmp_path, sim_path)

    print(f"\nSuccessfully graded {graded_count}/{len(games)} final games for '{url_date}' -> {sim_path}!")

if __name__ == "__main__":
    grade_games(url_date)
