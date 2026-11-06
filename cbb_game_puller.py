from cbb_sim_functional import run_simulation, calculate_league_averages, assign_full_team_stats
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
from openpyxl import load_workbook

# Step 1: Fetch the webpage content
url = input("Enter url: ")
response = requests.get(url)
response.raise_for_status()

# Step 2: Parse the HTML content
soup = BeautifulSoup(response.text, 'html.parser')

# Step 3: Find all game containers
game_containers = soup.find_all('div', class_='pbox')

if not game_containers:
    print("No game containers found. Check the HTML structure.")
    exit()

# Step 4: Extract data from each table
games_data = []
matchups = []

for container in game_containers:
    table = container.find('table')
    if not table:
        print("No table found in a game container.")
        continue

    rows = table.find_all('tr')
    teams = []
    for row in rows:
        if row.find('th'):
            continue

        cells = row.find_all('td')
        if len(cells) < 6:
            continue

        team_name = re.sub(r'\d+', '', cells[1].text.strip())
        teams.append(team_name)

        games_data.append({
            "Team Name": team_name,
            "Score": cells[2].text.strip(),
            "Spread": cells[3].text.strip(),
            "Win Probability": cells[4].text.strip(),
            "Confidence": cells[5].text.strip(),
        })

        if len(teams) == 2:
            matchups.append((teams[0], teams[1]))
            teams = []
            games_data.append({"Team Name": "", "Score": "", "Spread": "", "Win Probability": "", "Confidence": ""})

# Step 5: Convert data to a DataFrame
df = pd.DataFrame(games_data)

# Step 6: Load simulation dataset
stats_file = "2024-25 Men's College Basketball Stats.xlsx"
df_stats = pd.read_excel(stats_file, header=1)
league_averages = calculate_league_averages(df_stats)

# Step 7: Run simulations for each matchup
sim_results = []
for team1, team2 in matchups:
    result = run_simulation(team1, team2, df_stats, league_averages)
    if result:
        sim_results.append(result)

# Step 8: Save results to Excel
output_file = "cbb_game_data.xlsx"
sheet_name = input("Enter date: ")

sim_df = pd.DataFrame(sim_results)

try:
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        sim_df.to_excel(writer, index=False, sheet_name=f"{sheet_name}_Sim")
except FileNotFoundError:
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        sim_df.to_excel(writer, index=False, sheet_name=f"{sheet_name}_Sim")

print(f"Data and simulations saved to '{output_file}'")
