import asyncio
from pyppeteer import launch
import pandas as pd

async def fetch_table():
    # Launch a headless browser
    browser = await launch()
    page = await browser.newPage()

    # Navigate to the URL
    await page.goto("https://www.ncaa.com/rankings/basketball-men/d1/ncaa-mens-basketball-net-rankings")

    # Wait for the table to load
    await page.waitForSelector('table.sticky.tablesorter')

    # Extract the table HTML
    table_html = await page.evaluate('document.querySelector("table.sticky.tablesorter").outerHTML')

    # Close the browser
    await browser.close()

    # Parse the table HTML with pandas
    df = pd.read_html(table_html)[0]

    # Save to Excel
    df.to_excel("NCAA_NET_Rankings.xlsx", index=False)
    print("Data saved to NCAA_NET_Rankings.xlsx")

# Run the async function
asyncio.get_event_loop().run_until_complete(fetch_table())
