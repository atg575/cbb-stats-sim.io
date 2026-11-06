from cbb_sim_live_moreyz import simulate_game_with_rebounds_and_adjustments as adj_m_sim, calculate_league_averages, calculate_relative_strengths as adj_m_strengths, assign_full_team_stats, load_dataset
from cbb_sim_consensus_moreyz import simulate_game_with_rebounds_and_adjustments as con_m_sim, calculate_relative_strengths as con_m_strengths
from cbb_sim_live_log5 import simulate_game_with_rebounds_and_adjustments as adj_sim, calculate_relative_strengths as adj_strengths
from cbb_sim_consensus_log5 import simulate_game_with_rebounds_and_adjustments as con_sim, calculate_relative_strengths as con_strengths
from cbb_sim_adj_eff import simulate_game_with_rebounds_and_adjustments as old_sim, calculate_relative_strengths as old_strengths
from cbb_sim_functional import simulate_game_with_rebounds_and_adjustments as reg_sim, calculate_relative_strengths as reg_strengths
from game_extractor import extract_games
import csv
import os
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
import sys

file_path = "2026-27 Men's College Basketball Stats_code.xlsx"
dataset, nat_avgs = load_dataset(file_path)
league_averages = calculate_league_averages(dataset)
game_num = 0
url_date = sys.argv[1] if len(sys.argv) > 1 else datetime.now(ZoneInfo("America/New_York")).strftime("%Y-%m-%d")
url = "https://www.warrennolan.com/basketball/2027/scores-live?type=All%20Games&date=" + url_date

schedule = extract_games(url)
adj_m_win_pcts = []
adj_m_avg_scores = []
con_m_win_pcts = []
con_m_avg_scores = []
old_win_pcts = []
old_avg_scores = []
adj_win_pcts = []
adj_avg_scores = []
con_win_pcts = []
con_avg_scores = []
reg_win_pcts = []
reg_avg_scores = []


for game in schedule:
    game_num += 1
    teams = [game_num]
    for team in game:
        teams.append(team)

    dataset['Warren Name'] = dataset['Warren Name'].str.strip()
    dataset['Warren Name'] = dataset['Warren Name'].str.lower()  # Standardize casing

    [away_score, home_score, away_wins, home_wins, ot_games, double_ot_games] = [0,0,0,0,0,0]
        
    team1_data = assign_full_team_stats(dataset, teams[1])
    team2_data = assign_full_team_stats(dataset, teams[2])

    team1_adj_m_strengths = adj_m_strengths(team1_data['team'], team2_data['opponent'], league_averages)
    team2_adj_m_strengths = adj_m_strengths(team2_data['team'], team1_data['opponent'], league_averages)
    team1_con_m_strengths = con_m_strengths(team1_data['team'], team2_data['opponent'], league_averages)
    team2_con_m_strengths = con_m_strengths(team2_data['team'], team1_data['opponent'], league_averages)
    team1_adj_strengths = adj_strengths(team1_data['team'], team2_data['opponent'], league_averages)
    team2_adj_strengths = adj_strengths(team2_data['team'], team1_data['opponent'], league_averages)
    team1_con_strengths = con_strengths(team1_data['team'], team2_data['opponent'], league_averages)
    team2_con_strengths = con_strengths(team2_data['team'], team1_data['opponent'], league_averages)
    team1_old_strengths = old_strengths(team1_data['team'], team2_data['opponent'], league_averages)
    team2_old_strengths = old_strengths(team2_data['team'], team1_data['opponent'], league_averages)
    team1_reg_strengths = reg_strengths(team1_data['team'], team2_data['opponent'], league_averages)
    team2_reg_strengths = reg_strengths(team2_data['team'], team1_data['opponent'], league_averages)
    
    con_m_team_strengths = [team1_con_m_strengths, team2_con_m_strengths]
    adj_m_team_strengths = [team1_adj_m_strengths, team2_adj_m_strengths]
    old_team_strengths = [team1_old_strengths, team2_old_strengths]
    con_team_strengths = [team1_con_strengths, team2_con_strengths]
    adj_team_strengths = [team1_adj_strengths, team2_adj_strengths]
    reg_team_strengths = [team1_reg_strengths, team2_reg_strengths]

    # Safe tempo extraction with fallbacks to prevent NaN crashes
    t1_live_tempo = team1_data['team'].get('Live Tempo')
    t2_live_tempo = team2_data['team'].get('Live Tempo')
    if pd.isna(t1_live_tempo): t1_live_tempo = 68.0
    if pd.isna(t2_live_tempo): t2_live_tempo = 68.0

    t1_con_tempo = team1_data['team'].get('Consensus Tempo')
    t2_con_tempo = team2_data['team'].get('Consensus Tempo')
    if pd.isna(t1_con_tempo): t1_con_tempo = 68.0
    if pd.isna(t2_con_tempo): t2_con_tempo = 68.0

    t1_adj_tempo = team1_data['team'].get('Adj. Tempo')
    t2_adj_tempo = team2_data['team'].get('Adj. Tempo')
    if pd.isna(t1_adj_tempo): t1_adj_tempo = 68.0
    if pd.isna(t2_adj_tempo): t2_adj_tempo = 68.0

    t1_pace = team1_data['team'].get('Pace')
    t2_pace = team2_data['team'].get('Pace')
    if pd.isna(t1_pace): t1_pace = 68.0
    if pd.isna(t2_pace): t2_pace = 68.0

    adj_m_num_possessions = int(round((t1_live_tempo + t2_live_tempo) / 2))
    con_m_num_possessions = int(round((t1_con_tempo + t2_con_tempo) / 2))
    adj_num_possessions = int(round((t1_live_tempo + t2_live_tempo) / 2))
    con_num_possessions = int(round((t1_con_tempo + t2_con_tempo) / 2))
    old_num_possessions = int(round((t1_adj_tempo + t2_adj_tempo) / 2))
    reg_num_possessions = int(round((t1_pace + t2_pace) / 2))

    adj_m_away_score = 0
    adj_m_home_score = 0
    adj_m_away_wins = 0
    adj_m_home_wins = 0
    adj_m_ot_games = 0
    adj_m_double_ot_games = 0    

    # Simulate 25000 adjusted model games
    for _ in range(0,24999):
        
        game_result = adj_m_sim(
        team1_data, team2_data, league_averages, adj_num_possessions, adj_team_strengths
        )

        if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
            adj_m_ot_games += 1
            game_result = adj_m_sim(
            team1_data, team2_data, league_averages, int(adj_num_possessions/8), adj_team_strengths)
            if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
                adj_m_double_ot_games += 1
                game_result = adj_m_sim(
                team1_data, team2_data, league_averages, int(adj_num_possessions/8), adj_team_strengths
                )

        if game_result['Team 1']['Score'] > game_result['Team 2']['Score']:
            adj_m_away_wins += 1
        elif game_result['Team 1']['Score'] < game_result['Team 2']['Score']:
            adj_m_home_wins += 1

        adj_m_away_score += game_result['Team 1']['Score']
        adj_m_home_score += game_result['Team 2']['Score']

    adj_m_away_win_pct = round(adj_m_away_wins/25000, 4)
    adj_m_home_win_pct = round(adj_m_home_wins/25000, 4)
        
    adj_m_final_away_score = int(adj_m_away_score/25000)
    adj_m_final_home_score = int(adj_m_home_score/25000)

    # Change these two lines to append to the adj_m_ lists
    adj_m_win_pcts.append((adj_m_away_win_pct, adj_m_home_win_pct))
    adj_m_avg_scores.append((adj_m_final_away_score, adj_m_final_home_score))
    
    con_m_away_score = 0
    con_m_home_score = 0
    con_m_away_wins = 0
    con_m_home_wins = 0
    con_m_ot_games = 0
    con_m_double_ot_games = 0   

    # Simulate 25000 consensus model games
    for _ in range(0,24999):
        
        game_result = con_m_sim(
        team1_data, team2_data, league_averages, con_num_possessions, con_team_strengths
        )

        if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
            con_m_ot_games += 1
            game_result = con_m_sim(
            team1_data, team2_data, league_averages, int(con_num_possessions/8), con_team_strengths)
            if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
                con_m_double_ot_games += 1
                game_result = con_m_sim(
                team1_data, team2_data, league_averages, int(con_num_possessions/8), con_team_strengths
                )

        if game_result['Team 1']['Score'] > game_result['Team 2']['Score']:
            con_m_away_wins += 1
        elif game_result['Team 1']['Score'] < game_result['Team 2']['Score']:
            con_m_home_wins += 1

        con_m_away_score += game_result['Team 1']['Score']
        con_m_home_score += game_result['Team 2']['Score']

    con_m_away_win_pct = round(con_m_away_wins/25000, 4)
    con_m_home_win_pct = round(con_m_home_wins/25000, 4)
        
    con_m_final_away_score = int(con_m_away_score/25000)
    con_m_final_home_score = int(con_m_home_score/25000)

    # Change these two lines to append to the con_m_ lists
    con_m_win_pcts.append((con_m_away_win_pct, con_m_home_win_pct))
    con_m_avg_scores.append((con_m_final_away_score, con_m_final_home_score)) 

    adj_away_score = 0
    adj_home_score = 0
    adj_away_wins = 0
    adj_home_wins = 0
    adj_ot_games = 0
    adj_double_ot_games = 0    

    # Simulate 25000 adjusted model games
    for _ in range(0,24999):
        
        game_result = adj_sim(
        team1_data, team2_data, league_averages, adj_num_possessions, adj_team_strengths
        )

        if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
            adj_ot_games += 1
            game_result = adj_sim(
            team1_data, team2_data, league_averages, int(adj_num_possessions/8), adj_team_strengths)
            if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
                adj_double_ot_games += 1
                game_result = adj_sim(
                team1_data, team2_data, league_averages, int(adj_num_possessions/8), adj_team_strengths
                )

        if game_result['Team 1']['Score'] > game_result['Team 2']['Score']:
            adj_away_wins += 1
        elif game_result['Team 1']['Score'] < game_result['Team 2']['Score']:
            adj_home_wins += 1

        adj_away_score += game_result['Team 1']['Score']
        adj_home_score += game_result['Team 2']['Score']

    adj_away_win_pct = round(adj_away_wins/25000, 4)
    adj_home_win_pct = round(adj_home_wins/25000, 4)
        
    adj_final_away_score = int(adj_away_score/25000)
    adj_final_home_score = int(adj_home_score/25000)

    adj_win_pcts.append((adj_away_win_pct,adj_home_win_pct))
    adj_avg_scores.append((adj_final_away_score,adj_final_home_score))
    
    con_away_score = 0
    con_home_score = 0
    con_away_wins = 0
    con_home_wins = 0
    con_ot_games = 0
    con_double_ot_games = 0   

    # Simulate 25000 consensus model games
    for _ in range(0,24999):
        
        game_result = con_sim(
        team1_data, team2_data, league_averages, con_num_possessions, con_team_strengths
        )

        if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
            con_ot_games += 1
            game_result = con_sim(
            team1_data, team2_data, league_averages, int(con_num_possessions/8), con_team_strengths)
            if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
                con_double_ot_games += 1
                game_result = con_sim(
                team1_data, team2_data, league_averages, int(con_num_possessions/8), con_team_strengths
                )

        if game_result['Team 1']['Score'] > game_result['Team 2']['Score']:
            con_away_wins += 1
        elif game_result['Team 1']['Score'] < game_result['Team 2']['Score']:
            con_home_wins += 1

        con_away_score += game_result['Team 1']['Score']
        con_home_score += game_result['Team 2']['Score']

    con_away_win_pct = round(con_away_wins/25000, 4)
    con_home_win_pct = round(con_home_wins/25000, 4)
        
    con_final_away_score = int(con_away_score/25000)
    con_final_home_score = int(con_home_score/25000)

    con_win_pcts.append((con_away_win_pct,con_home_win_pct))
    con_avg_scores.append((con_final_away_score,con_final_home_score)) 

    old_away_score = 0
    old_home_score = 0 
    old_away_wins = 0
    old_home_wins = 0
    old_ot_games = 0
    old_double_ot_games = 0

    # Simulate 25000 adjusted efficiency (old) model games
    for _ in range(0,24999):
 
        game_result = old_sim(
        team1_data, team2_data, league_averages, old_num_possessions, old_team_strengths
        )

        if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
            old_ot_games += 1
            game_result = old_sim(
            team1_data, team2_data, league_averages, int(old_num_possessions/8), old_team_strengths)
            if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
                old_double_ot_games += 1
                game_result = old_sim(
                team1_data, team2_data, league_averages, int(old_num_possessions/8), old_team_strengths
                )

        if game_result['Team 1']['Score'] > game_result['Team 2']['Score']:
            old_away_wins += 1
        elif game_result['Team 1']['Score'] < game_result['Team 2']['Score']:
            old_home_wins += 1

        old_away_score += game_result['Team 1']['Score']
        old_home_score += game_result['Team 2']['Score']

    old_away_win_pct = round(old_away_wins/25000, 4)
    old_home_win_pct = round(old_home_wins/25000, 4)
        
    old_final_away_score = int(old_away_score/25000)
    old_final_home_score = int(old_home_score/25000)
    
    old_win_pcts.append((old_away_win_pct,old_home_win_pct))
    old_avg_scores.append((old_final_away_score,old_final_home_score))

    reg_away_score = 0
    reg_home_score = 0
    reg_away_wins = 0
    reg_home_wins = 0
    ot_games = 0
    double_ot_games = 0

# Simulate 25000 regular model games
    for _ in range(0,24999):
 
        game_result = reg_sim(
        team1_data, team2_data, league_averages, reg_num_possessions, reg_team_strengths
        )

        if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
            ot_games += 1
            game_result = reg_sim(
            team1_data, team2_data, league_averages, int(reg_num_possessions/8), reg_team_strengths)
            if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
                double_ot_games += 1
                game_result = reg_sim(
                team1_data, team2_data, league_averages, int(reg_num_possessions/8), reg_team_strengths
                )

        if game_result['Team 1']['Score'] > game_result['Team 2']['Score']:
            reg_away_wins += 1
        elif game_result['Team 1']['Score'] < game_result['Team 2']['Score']:
            reg_home_wins += 1

        reg_away_score += game_result['Team 1']['Score']
        reg_home_score += game_result['Team 2']['Score']

    reg_away_win_pct = round(reg_away_wins/25000, 4)
    reg_home_win_pct = round(reg_home_wins/25000, 4)
        
    reg_final_away_score = int(reg_away_score/25000)
    reg_final_home_score = int(reg_home_score/25000)
    
    reg_win_pcts.append((reg_away_win_pct,reg_home_win_pct))
    reg_avg_scores.append((reg_final_away_score,reg_final_home_score))

    # Display results
    print(f'''
        {game_result['Team 1']['Name']} vs {game_result['Team 2']['Name']}
        ----------------
        LIVE MODEL:
        ----------------
        Win Possibility Percentage:
        {game_result['Team 1']['Name']} {adj_m_away_win_pct} - 
        {game_result['Team 2']['Name']} {adj_m_home_win_pct}
        
        Final Score:
        {game_result['Team 1']['Name']} {adj_m_final_away_score} - 
        {game_result['Team 2']['Name']} {adj_m_final_home_score}
        ------------------------------------------------------''')
    
headers = [
        'School', 
        'Live Morey Win %', 'Live Morey Score', 
        '', 
        'Consensus Morey Win %', 'Consensus Morey Score', 
        '', 
        'Live Log5 Win %', 'Live Log5 Score', 
        '', 
        'Consensus Log5 Win %', 'Consensus Log5 Score', 
        '', 
        'Adj Eff Win %', 'Adj Eff Score',
        '', 
        'Regular Win %', 'Regular Score'
    ]

    # 1. Collect all rows into a list
output_data = []

    # i is the index (0, 1, 2...), game_teams is the tuple ('uf', 'lsu')
for i, game_teams in enumerate(schedule):
        
        # Loop through the two teams in the current game (index 0 and 1)
    for team_idx in range(len(game_teams)):
            
            school_name = game_teams[team_idx]
            
            # Use 'i' to find the right game, 'team_idx' to find the right team stats
            adj_m_win_pct = adj_m_win_pcts[i][team_idx]
            adj_m_score = adj_m_avg_scores[i][team_idx]
            con_m_win_pct = con_m_win_pcts[i][team_idx]
            con_m_score = con_m_avg_scores[i][team_idx]
            adj_win_pct = adj_win_pcts[i][team_idx]
            adj_score = adj_avg_scores[i][team_idx]
            con_win_pct = con_win_pcts[i][team_idx]
            con_score = con_avg_scores[i][team_idx]
            old_win_pct = old_win_pcts[i][team_idx]
            old_score = old_avg_scores[i][team_idx]
            reg_win_pct = reg_win_pcts[i][team_idx]
            reg_score = reg_avg_scores[i][team_idx]

            # Empty strings ("") create the blank column cells separating the models
            row = [
                school_name, 
                adj_m_win_pct, adj_m_score, 
                "", 
                con_m_win_pct, con_m_score, 
                "", 
                adj_win_pct, adj_score, 
                "", 
                con_win_pct, con_score, 
                "", 
                old_win_pct, old_score,
                "", 
                reg_win_pct, reg_score
            ]
            output_data.append(row)
            
        # Add an empty row between matchups for readability
    empty_row = [""] * len(headers)
    output_data.append(empty_row)

# 2. Convert the collected data into a Pandas DataFrame
results_df = pd.DataFrame(output_data, columns=headers)

    # 3. Write to an existing Excel workbook using mode='a'
    # You can change this to file_path if you want it in your master workbook
results_file = "Batch_Simulation_Results.xlsx"
sheet_name = url_date # Uses the date you typed in as the sheet name

if os.path.exists(results_file):
    with pd.ExcelWriter(results_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            results_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"\nData replaced/added in {results_file} on sheet '{sheet_name}'.")
else:
    with pd.ExcelWriter(results_file, engine='openpyxl') as writer:
            results_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"\nCreated new file {results_file} with sheet '{sheet_name}'.")