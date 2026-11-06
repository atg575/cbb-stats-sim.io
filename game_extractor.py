import requests
from bs4 import BeautifulSoup
# CORRECT (safely allows automated arguments or defaults):
import sys
from datetime import datetime
from zoneinfo import ZoneInfo

if len(sys.argv) > 1:
    url_date = sys.argv[1]
else:
    url_date = datetime.now(ZoneInfo("America/New_York")).strftime("%Y-%m-%d")


url = "https://www.warrennolan.com/basketball/2026/scores-live?type=All%20Games&date="+url_date

def extract_games(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # 1. Find all tables (each game is usually its own small table or section)
        tables = soup.find_all('table')

        games = []
        for table in tables:
            # 2. Inside each table, look ONLY for 'a' tags (links)
            links = table.find_all('a')
            
            current_game_teams = []
            
            for link in links:
                href = link.get('href', '')
                
                # 3. Filter: Only keep links that go to a team schedule page
                # The structure is usually /basketball/2026/schedule/Team-Name
                if '/schedule/' in href:
                    team_name = link.get_text(strip=True)
                    team_name = team_name.lower()
                    current_game_teams.append(team_name)

            # 4. Print clean results
            if len(current_game_teams) == 2:
                games.append((current_game_teams[0], current_game_teams[1]))
            elif len(current_game_teams) == 1:
                pass

    else:
        print("Failed to retrieve page")
    return games

games = extract_games(url)
print(games)