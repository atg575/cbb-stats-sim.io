import pandas as pd
import random

# Calculate league averages
def calculate_league_averages(df):
    """
    Calculate league averages for key stats.
    """
    averages = {
        'ORtg': df['ORtg'].mean(),
        'DRtg': df['Opp ORtg'].mean(),
        'TS%': df['TS%'].mean(),
        'Opp TS%': df['Opp TS%'].mean(),
        'Pace': df['Pace'].mean(),
        'FT%': df['FT%'].mean(),
        'eFG%' : df['eFG%'].mean(), 
        'Opp eFG%' : df['Opp eFG%'].mean(),
        'STL%' : df['STL%'].mean(),
        'Opp STL%' : df['Opp STL%'].mean(),
        'BLK%' : df['BLK%'].mean(),
        'Adj. Net FF' : df['Adj. Net FF'].mean(),
        'Adj. ORtg': df['Adj. ORtg'].mean(),
        'Adj. Tempo': df['Adj. Tempo'].mean(),
        'Adj. DRtg': df['Adj. DRtg'].mean()
    }
    return averages

# Calculate relative strengths
def calculate_relative_strengths(team_stats, opponent_stats, league_averages):
    """
    Calculate relative strengths for a team compared to the league average.
    """
    relative_strengths = {
        'Offensive Advantage': (team_stats['ORtg'] / league_averages['ORtg'])*(1+0.001*(team_stats['SRS']-opponent_stats['SRS'])),
        'Defensive Advantage': (league_averages['DRtg'] / opponent_stats['Opp ORtg'])/(1+0.001*(opponent_stats['SRS']-team_stats['SRS'])),
        'Shooting Advantage': (team_stats['eFG%'] / league_averages['eFG%'])*(1+0.001*(team_stats['SRS']-opponent_stats['SRS'])),
        'Opponent Shooting Defense': (league_averages['Opp eFG%']/opponent_stats['Opp eFG%'])/(1+0.001*(opponent_stats['SRS']-team_stats['SRS'])),
        'Pace Adjustment': (team_stats['Pace'] / league_averages['Pace'])*(1+0.001*(team_stats['SRS']-opponent_stats['SRS'])),
    }
    return relative_strengths

# Simulate a possession
def simulate_possession_with_balanced_probs(team_stats, opponent_stats, league_averages, relative_strengths):
    """
    Simulate a possession incorporating balanced probabilities.
    """
    two_attemp_prob = (team_stats['2PA'] / team_stats['Possessions'] +
                       (opponent_stats['Opp 2PA'] / opponent_stats['Opp Possessions'])) / 2
    two_attemp_prob *= relative_strengths['Pace Adjustment'] * relative_strengths['Offensive Advantage'] / relative_strengths['Defensive Advantage']

    three_attempt_prob = (team_stats['3PA'] / team_stats['Possessions'] +
                       (opponent_stats['Opp 3PA'] / opponent_stats['Opp Possessions'])) / 2
    three_attempt_prob *= relative_strengths['Pace Adjustment'] * relative_strengths['Offensive Advantage'] / relative_strengths['Defensive Advantage']
    
    fg_success_prob = ((team_stats['2P%']+ opponent_stats['Opp 2P%']) / 2) *relative_strengths['Shooting Advantage'] * relative_strengths['Opponent Shooting Defense']

    three_success_prob = ((team_stats['3P%'] + opponent_stats['Opp 3P%']) / 2) *relative_strengths['Shooting Advantage'] * relative_strengths['Opponent Shooting Defense']

    turnover_prob = ((team_stats['TOV%'] + opponent_stats['Opp TOV%']) / 2)

    offensive_rebound_prob = ((team_stats['ORB%'] + opponent_stats['Opp ORB%']) / 2)

    free_throw_prob = ((team_stats['FTr'] + opponent_stats['Opp FTr']) / 2) 

    foul_prob = (team_stats['PF'] + opponent_stats['Opp PF'])/(4*team_stats['Possessions'])

    total_prob = turnover_prob + three_attempt_prob + two_attemp_prob + foul_prob

    turnover_prob /= total_prob
    three_attempt_prob /= total_prob
    two_attemp_prob /= total_prob
    foul_prob /= total_prob

    # Determine outcome
    rand = random.random()
    if rand < turnover_prob:
        return "Turnover"
    elif rand < turnover_prob + three_attempt_prob:
            rand = random.random()    
            if rand < three_success_prob:
                if random.random() < (free_throw_prob/20):
                    return "Made 3pt and 1"
                return "Made 3pt"
            rand = random.random() 
            if rand < (free_throw_prob/10):
                return "3 Free Throws"
            else:
                rand = random.random()
                if rand < offensive_rebound_prob:
                    return "Offensive Rebound"
                else:
                    return "Defensive Rebound"
    elif rand < turnover_prob + three_attempt_prob + two_attemp_prob:
            rand = random.random()            
            if rand < fg_success_prob:
                if random.random() < free_throw_prob/5:
                    return "Made 2pt and 1"
                return "Made 2pt" 
            rand = random.random()           
            if rand < free_throw_prob:
                return "2 Free Throws"
            else:
                rand = random.random()
                if rand < offensive_rebound_prob:
                    return "Offensive Rebound"
                else:
                    return "Defensive Rebound"
    elif rand < turnover_prob + three_attempt_prob + two_attemp_prob + foul_prob:
        return "Personal Foul"

# Assign team stats
def assign_full_team_stats(df, team_name):
    if team_name not in df['School'].values and team_name not in df['Warren Name'].values:
        raise ValueError(f"Team '{team_name}' not found in the dataset.")
    if team_name in df['School'].values:
        team_data = df[df['School'] == team_name].iloc[0]
    elif team_name in df['Warren Name'].values:
        team_data = df[df['Warren Name'] == team_name].iloc[0]
# Calculate the average number of possessions
    team_stats = {
        'Kenpom Name': team_data['Kenpom Name'],
        'School': team_name,
        'Possessions': team_data['Possessions'],  # Ensure this matches your dataset
        'Pace': team_data['Pace'],
        'FG': team_data['FG'],
        'FGA': team_data['FGA'],
        'FG%': team_data['FG%'],
        '2P': team_data['2P'],
        '2PA': team_data['2PA'],
        '2P%': team_data['2P%'],        
        '3P': team_data['3P'],
        '3PA': team_data['3PA'],
        '3P%': team_data['3P%'],
        'FT%': team_data['FT%'],
        'ORB': team_data['ORB'],
        'TRB': team_data['TRB'],
        'AST': team_data['AST'],
        'STL': team_data['STL'],
        'BLK': team_data['BLK'],
        'TOV': team_data['TOV'],
        'PF': team_data['PF'],
        'ORtg': team_data['ORtg'],
        'FTr': team_data['FTr'],
        '3PAr': team_data['3PAr'],
        'TS%': team_data['TS%'],
        'TRB%': team_data['TRB%'],
        '2C%': team_data['2C%'],
        'AST%': team_data['AST%'],
        'STL%': team_data['STL%'],
        'BLK%': team_data['BLK%'],
        'eFG%': team_data['eFG%'],
        'TOV%': team_data['TOV%'],
        'ORB%': team_data['ORB%'],
        'FT/FGA': team_data['FT/FGA'],
        'PF': team_data['Opp PF'],
        'Adj. Net FF': team_data['Adj. Net FF'],
        'SRS': team_data['SRS'],
    }
    opponent_stats = {
        'Opp PF': team_data['PF'],
        'Opp FG': team_data['Opp FG'],
        'Opp FGA': team_data['Opp FGA'],
        'Opp FG%': team_data['Opp FG%'],
        'Opp 2P': team_data['Opp 2P'],
        'Opp 2PA': team_data['Opp 2PA'],
        'Opp 2P%': team_data['Opp 2P%'],        
        'Opp 3P': team_data['Opp 3P'],
        'Opp 3PA': team_data['Opp 3PA'],
        'Opp 3P%': team_data['Opp 3P%'],
        'Opp ORB': team_data['Opp ORB'],
        'Opp TRB': team_data['Opp TRB'],
        'Opp AST': team_data['Opp AST'],
        'Opp STL': team_data['Opp STL'],
        'Opp BLK': team_data['Opp BLK'],
        'Opp TOV': team_data['Opp TOV'],
        'Opp PF': team_data['Opp PF'],
        'Opp ORtg': team_data['Opp ORtg'],
        'Opp FTr': team_data['Opp FTr'],
        'Opp 3PAr': team_data['Opp 3PAr'],
        'Opp TS%': team_data['Opp TS%'],
        'Opp TRB%': team_data['Opp TRB%'],
        'Opp 2C%': team_data['Opp 2C%'],
        'Opp AST%': team_data['Opp AST%'],
        'Opp STL%': team_data['Opp STL%'],
        'Opp BLK%': team_data['Opp BLK%'],
        'Opp eFG%': team_data['Opp eFG%'],
        'Opp TOV%': team_data['Opp TOV%'],
        'Opp ORB%': team_data['Opp ORB%'],
        'Opp FT/FGA': team_data['Opp FT/FGA'],
        'Opp Possessions': team_data['Opp Possessions'],
        'Adj. Net FF': team_data['Adj. Net FF'],
        'Adj. ORtg': team_data['Adj. ORtg'],
        'Adj. DRtg': team_data['Adj. DRtg'],
        'Adj. Net Rtg': team_data['Adj. Net Rtg'],
        'Adj. Tempo': team_data['Adj. Tempo'],
        'SRS': team_data['SRS']
    }
    
    return {'team': team_stats, 'opponent': opponent_stats}

# Simulate a game
def simulate_game_with_rebounds_and_adjustments(
    team1_data, team2_data, league_averages, num_possessions, team_strengths
):
    """
    Simulate a full basketball game incorporating detailed rebound logging and adjusted free throw probabilities.
    """
    team1_score = 0
    team2_score = 0
    team1_fouls = 0
    team2_fouls = 0
    team1_strengths, team2_strengths, = team_strengths
    team1_log = []
    team2_log = []

    # After initializing logs and scores
    #print("Initial logs:", team1_log, team2_log)
    #print("Initial scores:", team1_score, team2_score)

    for _ in range(num_possessions):
        # Team 1 possession
        if _ == num_possessions/2:
            team1_fouls = 0
            team2_fouls = 0
        while True:
            team1_outcome = simulate_possession_with_balanced_probs(
            team1_data['team'], team2_data['opponent'], league_averages, team1_strengths
)
            #print(f"Team 1 outcome: {team1_outcome}")
            if team1_outcome == "Personal Foul":
                team2_fouls += 1
                if team2_fouls >= 7 and team2_fouls < 10:
                    team1_outcome = "1 and 1"   
                elif team2_fouls >= 10:
                    team1_outcome = "2 Free Throws"   
                else:
                    continue
                if team1_outcome == "1 and 1":
                    points = 0
                    for ft in range(0,1):
                        if random.random() < team1_data['team']['FT%']:
                            points += 1
                        else:
                            break
                    team1_score += points
                    team1_log.append({'Outcome': team1_outcome, 'Points': points})
                    if random.random() < ((team1_data['team']['ORB%']+ team2_data['opponent']['Opp ORB%'])/5):
                        team1_outcome = "Offensive Rebound"
                        team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                        continue
                    else:
                        team2_outcome = "Defensive Rebound"
                        team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                        break
                elif team1_outcome == "2 Free Throws":
                    points = 0
                    for ft in range(0,1):
                        points = 1 if random.random() < team1_data['team']['FT%'] else 0
                        team1_score += points
                    team1_log.append({'Outcome': team1_outcome, 'Points': points})
                    if ft == 1 and points == 0:
                        if random.random() < ((team1_data['team']['ORB%']+ team2_data['opponent']['Opp ORB%'])/5):
                            team1_outcome = "Offensive Rebound"
                            team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                            continue
                        else:
                            team1_outcome = "Defensive Rebound"
                            team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                            break  
                elif team1_outcome == "3 Free Throws":
                    points = 0
                    for ft in range(0,2):
                        points = 1 if random.random() < team1_data['team']['FT%'] else 0
                        team1_score += points
                    team1_log.append({'Outcome': team1_outcome, 'Points': points})
                    if ft == 2 and points == 0:
                        if random.random() < ((team1_data['team']['ORB%']+ team2_data['opponent']['Opp ORB%'])/5):
                            team1_outcome = "Offensive Rebound"
                            team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                            continue  
                        else:
                            team1_outcome = "Defensive Rebound"
                            team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                            break                    
            elif team1_outcome == "Made 2pt" or team1_outcome == "Made 2pt and 1":
                points = 2
                team1_score += points
                team1_log.append({'Outcome': team1_outcome, 'Points': team1_score})
                if team1_outcome == "Made 2pt and 1":
                    team1_fouls += 1
                    if random.random() < team1_data['team']['FT%']:
                        points += 1
                        team1_score += points
                        team1_log.append({'Outcome': team1_outcome, 'Points': team1_score})                        
                    elif random.random() < ((team1_data['team']['ORB%']+ team2_data['opponent']['Opp ORB%'])/5):
                        team1_outcome = "Offensive Rebound"
                        team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                        continue
                    else:
                        team1_outcome = "Defensive Rebound"
                        team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                        break                    
            elif team1_outcome == "Made 3pt" or team1_outcome == "Made 3pt and 1":
                points = 3
                team1_score += points
                team1_log.append({'Outcome': team1_outcome, 'Points': team1_score})
                if team1_outcome == "Made 3pt and 1":
                    team1_fouls += 1
                    if random.random() < team1_data['team']['FT%']:
                        points += 1
                        team1_score += points
                        team1_log.append({'Outcome': team1_outcome, 'Points': team1_score})                        
                    elif random.random() < ((team1_data['team']['ORB%']+ team2_data['opponent']['Opp ORB%'])/5):
                        team1_outcome = "Offensive Rebound"
                        team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                        continue
                    else:
                        team1_outcome = "Defensive Rebound"
                        team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                        break                         
            elif team1_outcome == "2 Free Throws":
                team2_fouls += 1
                points = 0
                for ft in range(0,1):
                    points = 1 if random.random() < team1_data['team']['FT%'] else 0
                    team1_score += points
                team1_log.append({'Outcome': team1_outcome, 'Points': points})
                if ft == 1 and points == 0:
                    if random.random() < ((team1_data['team']['ORB%']+ team2_data['opponent']['Opp ORB%'])/5):
                        team1_outcome = "Offensive Rebound"
                        team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                        continue
                    else:
                        team1_outcome = "Defensive Rebound"
                        team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                        break  
            elif team1_outcome == "3 Free Throws":
                team2_fouls += 1
                points = 0
                for ft in range(0,2):
                    points = 1 if random.random() < team1_data['team']['FT%'] else 0
                    team1_score += points
                team1_log.append({'Outcome': team1_outcome, 'Points': points})
                if ft == 2 and points == 0:
                    if random.random() < ((team1_data['team']['ORB%']+ team2_data['opponent']['Opp ORB%'])/5):
                        team1_outcome = "Offensive Rebound"
                        team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                        continue  
                    else:
                        team1_outcome = "Defensive Rebound"
                        team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                        break                         
            elif team1_outcome == "Offensive Rebound":
                team1_log.append({'Outcome': team1_outcome, 'Points': 0})
                continue  
            else:
                team1_log.append({'Outcome': "Defensive Rebound", 'Points': 0})
                break
            break
        # Team 2 possession
        while True:
            team2_outcome = simulate_possession_with_balanced_probs(
            team2_data['team'], team1_data['opponent'], league_averages, team2_strengths
            )
            #print(f"Team 2 outcome: {team2_outcome}")
            if team2_outcome == "Personal Foul":
                team1_fouls += 1
                if team1_fouls >= 7 and team1_fouls < 10:
                    team2_outcome = "1 and 1"   
                if team1_fouls >= 10:
                    team2_outcome = "2 Free Throws" 
                else:
                    continue                      
                if team2_outcome == "1 and 1":
                    points = 0
                    for ft in range(0,1):
                        if random.random() < team2_data['team']['FT%']:
                            points += 1
                        else:
                            break
                    team2_score += points
                    team2_log.append({'Outcome': team2_outcome, 'Points': points})
                    if random.random() < ((team2_data['team']['ORB%']+ team1_data['opponent']['Opp ORB%'])/5):
                        team2_outcome = "Offensive Rebound"
                        team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                        continue
                    else:
                        team2_outcome = "Defensive Rebound"
                        team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                        break
                elif team2_outcome == "2 Free Throws":
                    points = 0
                    for ft in range(0,1):
                        points += 1 if random.random() < team2_data['team']['FT%'] else 0
                        team2_score += points
                    team2_log.append({'Outcome': team2_outcome, 'Points': points})
                    if ft == 1 and points == 0:
                        if random.random() < ((team2_data['team']['ORB%']+ team1_data['opponent']['Opp ORB%'])/5):
                            team2_outcome = "Offensive Rebound"
                            team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                            continue
                        else:
                            team2_outcome = "Defensive Rebound"
                            team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                            break       
                elif team2_outcome == "3 Free Throws":
                    points = 0
                    for ft in range(0,2):
                        points = 1 if random.random() < team2_data['team']['FT%'] else 0
                        team2_score += points
                    team2_log.append({'Outcome': team2_outcome, 'Points': points})
                    if ft == 2 and points == 0:
                        if random.random() < ((team2_data['team']['ORB%']+ team1_data['opponent']['Opp ORB%'])/5):
                            team2_outcome = "Offensive Rebound"
                            team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                            continue
                        else:
                            team2_outcome = "Defensive Rebound"
                            team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                            break       
                elif team2_outcome == "Offensive Rebound":
                    team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                    continue 
                else:
                    team1_log.append({'Outcome': "Defensive Rebound", 'Points': 0})
                    break
                break                         
            elif team2_outcome == "Made 2pt" or team2_outcome == "Made 2pt and 1":
                points = 2
                team2_score += points
                team2_log.append({'Outcome': team2_outcome, 'Points': points})                 
                if team2_outcome == "Made 2pt and 1":
                    team1_fouls += 1
                    if random.random() < team2_data['team']['FT%']:
                        points += 1
                        team2_score += points
                        team2_log.append({'Outcome': team2_outcome, 'Points': team2_score})                        
                    elif random.random() < ((team2_data['team']['ORB%']+ team1_data['opponent']['Opp ORB%'])/5):
                        team2_outcome = "Offensive Rebound"
                        team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                        continue
                    else:
                        team2_outcome = "Defensive Rebound"
                        team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                        break  
            elif team2_outcome == "Made 3pt" or team2_outcome == "Made 3pt and 1":
                points = 3
                team2_score += points
                team2_log.append({'Outcome': team2_outcome, 'Points': points})                  
                if team2_outcome == "Made 3pt and 1":
                    team1_fouls += 1
                    if random.random() < team2_data['team']['FT%']:
                        points += 1
                        team2_score += points
                        team2_log.append({'Outcome': team2_outcome, 'Points': team2_score})                        
                    elif random.random() < ((team2_data['team']['ORB%']+ team1_data['opponent']['Opp ORB%'])/5):
                        team2_outcome = "Offensive Rebound"
                        team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                        continue
                    else:
                        team2_outcome = "Defensive Rebound"
                        team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                        break
            elif team2_outcome == "2 Free Throws":
                team1_fouls += 1
                points = 0
                for ft in range(0,1):
                    points += 1 if random.random() < team2_data['team']['FT%'] else 0
                    team2_score += points
                team2_log.append({'Outcome': team2_outcome, 'Points': points})
                if ft == 1 and points == 0:
                    if random.random() < ((team2_data['team']['ORB%']+ team1_data['opponent']['Opp ORB%'])/5):
                        team2_outcome = "Offensive Rebound"
                        team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                        continue
                    else:
                        team2_outcome = "Defensive Rebound"
                        team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                        break       
            elif team2_outcome == "3 Free Throws":
                team1_fouls += 1
                points = 0
                for ft in range(0,2):
                    points = 1 if random.random() < team2_data['team']['FT%'] else 0
                    team2_score += points
                team2_log.append({'Outcome': team2_outcome, 'Points': points})
                if ft == 2 and points == 0:
                    if random.random() < ((team2_data['team']['ORB%']+ team1_data['opponent']['Opp ORB%'])/5):
                        team2_outcome = "Offensive Rebound"
                        team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                        continue
                    else:
                        team2_outcome = "Defensive Rebound"
                        team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                        break       
            elif team2_outcome == "Offensive Rebound":
                team2_log.append({'Outcome': team2_outcome, 'Points': 0})
                continue 
            else:
                team1_log.append({'Outcome': "Defensive Rebound", 'Points': 0})
                break
            break
        
    # Before returning the result
    #print("Final scores:", team1_score, team2_score)

    return {
        'Team 1': {'Name': team1_data['team']['Kenpom Name'], 'Score': team1_score, 'Log': team1_log},
        'Team 2': {'Name': team2_data['team']['Kenpom Name'], 'Score': team2_score, 'Log': team2_log},
    }

# Correct file path
file_path = r"/Users/alexvazquez/Coding/cbb_stats/2026-27 Men's College Basketball Stats_code.xlsx"

# Load the dataset
try:
    # Adjust header row based on your dataset's structure
    df = pd.read_excel(file_path, sheet_name='Worksheet', header=1)  # Use header=0 if the first row contains column names
except FileNotFoundError:
    raise FileNotFoundError(f"File not found. Ensure the file exists at: {file_path}")

# Check and adjust columns
#print("Dataset Columns:", df.columns.tolist())

# Ensure relevant columns match the dataset
relevant_columns = [
    'Kenpom Name', 'School', 'Possessions', 'Pace', 'FG', 'FGA', 'FG%', '2P', '2PA', '2P%', '3P', '3PA', '3P%', 
    'FT%', 'ORB', 'TRB', 'AST', 'STL', 'BLK', 'TOV', 'PF', 'ORtg',
    'FTr', '3PAr', 'TS%', 'TRB%', 'AST%', 'STL%', 'BLK%', 'eFG%', 'TOV%', 
    'ORB%', 'FT/FGA', 'Opp FG', 'Opp FGA', 'Opp FG%', 'Opp 2P', 'Opp 2PA', 
    'Opp 2P%', 'Opp 3P', 'Opp 3PA', 'Opp 3P%', 'Opp ORB', 'Opp TRB', 'Opp AST', 
    'Opp STL', 'Opp BLK', 'Opp TOV', 'Opp PF', 'Opp ORtg', 'Opp FTr', 'Opp 3PAr', 
    'Opp TS%', 'Opp TRB%', 'Opp AST%', 'Opp STL%', 'Opp BLK%', 'Opp eFG%', 
    'Opp TOV%', 'Opp ORB%', 'Opp FT/FGA', 'Opp Possessions', 'Adj. Net FF','Adj. Net FF',
    'Adj. ORtg','Adj. DRtg', 'Adj. Net Rtg', 'Adj. Tempo', 'Net Score', 'SRS'
]

# Check if all columns exist
missing_columns = [col for col in relevant_columns if col not in df.columns]
if missing_columns:
    raise KeyError(f"The following columns are missing from the dataset: {missing_columns}")

# Extract relevant columns
df_relevant_with_opponent = df[relevant_columns].copy()

df_relevant_with_opponent.head()

# Calculate league averages
league_averages = calculate_league_averages(df_relevant_with_opponent)

#print("Columns in dataset:", df.columns.tolist())
away_log = [];
home_log = [];

if __name__ == "__main__":
# Assign team stats
    
    away_name = input("Enter Away team name: ").strip().title()
    home_name = input("Enter Home team name: ").strip().title()

    df_relevant_with_opponent['School'] = df_relevant_with_opponent['School'].str.strip()
    df_relevant_with_opponent['School'] = df_relevant_with_opponent['School'].str.title()  # Standardize casing

    away_score = 0
    home_score = 0

    away_wins = 0
    home_wins = 0
    ot_games = 0
    double_ot_games = 0

    team1_name = away_name
    team2_name = home_name
        
    team1_data = assign_full_team_stats(df_relevant_with_opponent, team1_name)
    team2_data = assign_full_team_stats(df_relevant_with_opponent, team2_name)

    team1_strengths = calculate_relative_strengths(team1_data['team'], team2_data['opponent'], league_averages)
    team2_strengths = calculate_relative_strengths(team2_data['team'], team1_data['opponent'], league_averages)
    team_strengths = [team1_strengths, team2_strengths]

    num_possessions = int(round((team1_data['team']['Possessions'] + team1_data['opponent']['Opp Possessions']+ team2_data['team']['Possessions']+team2_data['opponent']['Opp Possessions'])/2.1))
    
    # Simulate 10000 games
    for _ in range(0,24999):
        
        game_result = simulate_game_with_rebounds_and_adjustments(
        team1_data, team2_data, league_averages, num_possessions, team_strengths
        )

        if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
            ot_games += 1
            game_result = simulate_game_with_rebounds_and_adjustments(
            team1_data, team2_data, league_averages, int(num_possessions/8))
            if game_result['Team 1']['Score'] == game_result['Team 2']['Score']:
                double_ot_games += 1
                game_result = simulate_game_with_rebounds_and_adjustments(
                team1_data, team2_data, league_averages, int(num_possessions/8)
                )

        if game_result['Team 1']['Score'] > game_result['Team 2']['Score']:
            away_wins += 1
        elif game_result['Team 1']['Score'] < game_result['Team 2']['Score']:
            home_wins += 1

        away_score += game_result['Team 1']['Score']
        home_score += game_result['Team 2']['Score']
        
        away_log.append(away_score)
        home_log.append(home_score)

        team1_log = pd.DataFrame(game_result['Team 1']['Log'])
        team2_log = pd.DataFrame(game_result['Team 2']['Log'])

    away_win_pct = round(away_wins/250, 2)
    home_win_pct = round(home_wins/250, 2)
        
    final_away_score = int(away_score/25000)
    final_home_score = int(home_score/25000)

    # Display results
    print(f"Final Score: {game_result['Team 1']['Name']} {final_away_score} - "
        f"{game_result['Team 2']['Name']} {final_home_score}"
        f"\nWin Possibility Percentage: {game_result['Team 1']['Name']} {away_win_pct} - "
        f"{game_result['Team 2']['Name']} {home_win_pct}"
        f"\n{ot_games/100} OT Pct | {double_ot_games/100} 2OT Pct")


    # Extract logs
    team1_log.head(), team2_log.head()

    team_logs = {
        team1_name : away_log,
        team2_name: home_log
        }

    log_df = pd.DataFrame(team_logs)

    #log_df.to_csv(f"{team1_name}_vs_{team2_name}_game_log.csv")

#def run_simulation(team1_name, team2_name, df_stats, league_averages):
    #try:
       # team1_data = assign_full_team_stats(df_stats, team1_name)
        #team2_data = assign_full_team_stats(df_stats, team2_name)

        # Ensure possessions are available
       # num_possessions = int((team1_data['team'].get('Possessions', 70) + team2_data['team'].get('Possessions', 70)) / 3)

        # Simulate 1000 games
       # total_team1_score = 0
       # total_team2_score = 0
       # team1_wins = 0
       # team2_wins = 0

      #  for _ in range(1000):
      #      game_result = simulate_game_with_rebounds_and_adjustments(team1_data, team2_data, league_averages, num_possessions)
      #      total_team1_score += game_result['Team 1']['Score']
      #      total_team2_score += game_result['Team 2']['Score']

       #     if game_result['Team 1']['Score'] > game_result['Team 2']['Score']:
      #          team1_wins += 1
       #     else:
      #          team2_wins += 1

        # Calculate final scores and win probabilities
      #  team1_score = round(total_team1_score / 1000)
       # team2_score = round(total_team2_score / 1000)
       # team1_win_pct = round((team1_wins / 1000) * 100, 1)
       # team2_win_pct = round((team2_wins / 1000) * 100, 1)

      #  return {
      #      "Team 1": team1_name, 
       #     "Team 2": team2_name,
        #    "Simulated Score 1": team1_score, 
         #   "Simulated Score 2": team2_score,
          #  "Win Probability 1": f"{team1_win_pct}%",
           # "Win Probability 2": f"{team2_win_pct}%"
        #}
    
    #except Exception as e:
     #   print(f"Simulation failed for {team1_name} vs {team2_name}: {e}")
      #  return None


# Ensure no code runs when importing this file
#if __name__ == "__main__":
    # Example usage (ONLY runs when the script is executed directly)
   # team1_name = input("Enter Away team name: ").strip().title()
   # team2_name = input("Enter Home team name: ").strip().title()

   # stats_file = "2024-25 Men's College Basketball Stats.xlsx"
   # df_stats = pd.read_excel(stats_file, header=1)
   # league_averages = calculate_league_averages(df_stats)

   # result = run_simulation(team1_name, team2_name, df_stats, league_averages)
   # print(result)