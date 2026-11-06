from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time

# Step 1: Set up Selenium WebDriver
driver = webdriver.Chrome()  # Use the appropriate driver for your browser
url = "https://www.nytimes.com/athletic/college-basketball/odds/"
driver.get(url)

# Optional: Add a wait to ensure the page fully loads
time.sleep(5)

# Debug: Print the HTML to check if data is loaded
print("Rendered HTML:")
print(driver.page_source[:2000])  # Print first 2000 characters of HTML

# Step 2: Parse the rendered HTML content with BeautifulSoup
soup = BeautifulSoup(driver.page_source, 'html.parser')

# Step 3: Find parent containers for game data
parent_containers = soup.find_all('div', class_='game-odds-group_GameOddsBlocksContainer__CA4bj')  # Update class name if necessary

if not parent_containers:
    print("No parent containers found. Check the structure.")
    driver.quit()
    exit()

# Debug: Print the number of containers found
print(f"Found {len(parent_containers)} parent containers.")

# Step 4: Locate tables within the parent containers
games_data = []

for container in parent_containers:
    tables = container.find_all('table', class_='game-odds-block_OddsTable__r2cc_')  # Update class name if necessary
    for table in tables:
        rows = table.find('tbody').find_all('tr')
        if len(rows) < 1:
            continue

        # Extract team names from the first row
        team_cells = rows[0].find_all('td', class_='game-odds-block_TeamNameTableCell__NhNpl')
        if len(team_cells) != 2:
            continue  # Skip if there aren't exactly 2 teams

        team_1 = team_cells[0].text.strip()
        team_2 = team_cells[1].text.strip()

        # Extract odds (Spread, Total, Moneyline) from subsequent rows
        odds_cells = rows[0].find_all('td', class_='game-odds-block_OddsTableCell')
        if len(odds_cells) < 3:
            continue  # Skip if there aren't enough odds columns

        spread = odds_cells[0].text.strip()
        total = odds_cells[1].text.strip()
        moneyline = odds_cells[2].text.strip()

        # Add to games data
        games_data.append({
            "Team 1": team_1,
            "Team 2": team_2,
            "Spread": spread,
            "Total": total,
            "Moneyline": moneyline,
        })

# Step 5: Convert data to a DataFrame
if not games_data:
    print("No game data extracted. Check the structure and class names.")
    driver.quit()
    exit()

df = pd.DataFrame(games_data)

# Step 6: Save to Excel
output_file = "nyt_cbb_odds.xlsx"
sheet_name = "Odds Data"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name=sheet_name)

print(f"Data has been saved to the sheet '{sheet_name}' in {output_file}")

# Step 7: Quit the WebDriver
driver.quit()
