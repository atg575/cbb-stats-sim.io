from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import time

# Set up Selenium WebDriver
driver = webdriver.Chrome()  # Ensure ChromeDriver is installed
url = "https://www.vegasinsider.com/college-basketball/odds/las-vegas/"
driver.get(url)

# Wait for iframe to load
iframe = WebDriverWait(driver, 20).until(
    lambda d: d.find_element(By.TAG_NAME, "iframe")
)
iframe_src = iframe.get_attribute("src")
print("Iframe URL:", iframe_src)

# Navigate directly to iframe content
driver.get(iframe_src)
time.sleep(5)

# Parse the iframe's content
soup = BeautifulSoup(driver.page_source, 'html.parser')
print(soup.prettify())  # Debug HTML structure

# Adjust selectors based on observed structure
table = soup.find('table', {'class': 'odds-table-class'})  # Replace with actual class
if not table:
    print("Odds table not found. Verify the structure.")
    driver.quit()
    exit()

# Extract data
games_data = []
rows = table.find_all('tr')[1:]  # Skip header row
for row in rows:
    cols = row.find_all('td')
    if len(cols) > 1:
        games_data.append({
            'Team': cols[0].text.strip(),
            'Spread': cols[1].text.strip(),
            'Total': cols[2].text.strip(),
            'Moneyline': cols[3].text.strip(),
        })

# Convert to DataFrame
df = pd.DataFrame(games_data)

# Save to Excel
df.to_excel('college_basketball_odds.xlsx', index=False)
print("Data saved to 'college_basketball_odds.xlsx'")

driver.quit()
