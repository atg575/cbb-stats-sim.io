import csv

url_date = 'Day today'

schedule = [('uf', 'lsu'),('michigan','oklahoma'),('msu', 'louisville')]
win_percentages = [(0.74, 0.26), (0.55, 0.45), (0.82, 0.18)]
avg_score=   [(83, 78), (91, 90), (78, 65)]

headers = ['School', 'Win Percentage', 'Avg. Score']

with open(f"{url_date}-games.csv", 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(headers)
    
    # i is the index (0, 1, 2...), game_teams is the tuple ('uf', 'lsu')
    for i, game_teams in enumerate(schedule):
        
        # Loop through the two teams in the current game (index 0 and 1)
        for team_idx in range(len(game_teams)):
            
            school_name = game_teams[team_idx]
            
            # Use 'i' to find the right game, 'team_idx' to find the right team stats
            win_pct = win_percentages[i][team_idx]
            score = avg_score[i][team_idx]
            
            row = [school_name, win_pct, score]
            writer.writerow(row)
            
        # Optional: Add an empty row between matchups for readability
        writer.writerow([]) 

print("CSV created successfully.")