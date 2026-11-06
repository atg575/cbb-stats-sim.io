import requests
from bs4 import BeautifulSoup

# The URL for the A-Z list
url = "https://www.warrennolan.com/basketball/2026/teams-az"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

response = requests.get(url, headers=headers)

if response.status_code == 200:
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Use a set to automatically remove duplicates if a team appears twice
    teams = set()
    
    # Find all links on the page
    links = soup.find_all('a')
    
    for link in links:
        href = link.get('href', '')
        text = link.get_text(strip=True)
        
        # Filter logic:
        # 1. Must contain '/schedule/' (this isolates team pages)
        # 2. Must contain '/basketball/2026' (ensures we don't grab old years or other sports)
        # 3. Text must not be empty
        if '/basketball/2026/schedule/' in href and text:
            teams.add(text)
            
    # Convert set back to a list and sort A-Z
    sorted_teams = sorted(list(teams))
    
    # Print the results
    print(f"Found {len(sorted_teams)} teams:")
    for team in sorted_teams:
        print(team)

else:
    print(f"Failed to retrieve page: {response.status_code}")