import asyncio
import aiohttp
from lxml import html
import pandas as pd
import os
from urllib.parse import quote
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox
import requests
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, filename="scraper.log", format="%(asctime)s - %(message)s")

BASE_URL = "https://ppdvn.gov.vn/web/guest/ke-hoach-xuat-ban"
HEADER = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
}

def get_total_pages(title):
    try:
        query_encoded = quote(title)
        current_date = datetime.now().strftime("%d/%m/%Y")
        url = f"{BASE_URL}?query={query_encoded}&id_nxb=8&bat_dau=01%2F01%2F2004&ket_thuc={current_date}&p=1"
        response = requests.get(url, headers=HEADER)
        response.raise_for_status()

        tree = html.fromstring(response.text)
        pagination = tree.xpath("//div[@class='pagination']//a")

        if pagination:
            page_numbers = [
                int(link.get("href").split("p=")[-1])
                for link in pagination if "p=" in link.get("href", "")
            ]
            return max(page_numbers, default=1)
        else:
            print("No pagination found, defaulting to 1 page.")
            return 1

    except Exception as e:
        print(f"Error determining total pages: {e}")
        return 1

async def fetch_page(session, url):
    retries = 3
    for _ in range(retries):
        try:
            async with session.get(url, headers=HEADER) as response:
                response.raise_for_status()
                return await response.text()
        except Exception as e:
            print(f"Error fetching URL {url}: {e}")
            await asyncio.sleep(1)
    return None

def parse_page(html_content):
    try:
        tree = html.fromstring(html_content)
        rows = tree.xpath("//table[@cellpadding='0' and @cellspacing='0']/tbody/tr")
        data = []
        for row in rows:
            cells = row.xpath("td/text()")
            if len(cells) < 8:
                continue
            data.append({
                "STT": cells[0].strip(),
                "ISBN": cells[1].strip(),
                "Title": cells[2].strip(),
                "Author": cells[3].strip(),
                "Translator": cells[4].strip(),
                "Numbers of ordered copies": cells[5].strip(),
                "Self-publishing": cells[6].strip(),
                "Partner": cells[7].strip()
            })
        return data
    except Exception as e:
        print(f"Error parsing page: {e}")
        return []

async def scrape_pages(title, total_pages):
    query_encoded = quote(title)
    current_date = datetime.now().strftime("%d/%m/%Y")
    base_url = f"{BASE_URL}?query={query_encoded}&id_nxb=8&bat_dau=01%2F01%2F2004&ket_thuc={current_date}&p="

    async with aiohttp.ClientSession() as session:
        tasks = []
        for page in range(1, total_pages + 1):
            url = base_url + str(page)
            tasks.append(fetch_page(session, url))

        html_contents = await asyncio.gather(*tasks)
        all_data = []
        for content in html_contents:
            if content:
                all_data.extend(parse_page(content))
        return all_data

def save_to_excel(data, output, sheet_name):
    if data:
        df = pd.DataFrame(data)
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Data saved to {output} in sheet {sheet_name}")
    else:
        print("No data to save.")

def start_scraping(title, output):
    total_pages = get_total_pages(title)
    print(f"Total pages to scrape: {total_pages}")

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    data = loop.run_until_complete(scrape_pages(title, total_pages))

    save_to_excel(data, output, title)
    print("Scraping completed.")
    os.startfile(output)

def start_scraping_thread(title):
    output = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm")])
    if not output:
        return

    with ThreadPoolExecutor() as executor:
        executor.submit(start_scraping, title, output)

def open_github():
    webbrowser.open("https://github.com/DuckFard/Web-scraper-for-The-PPDA-of-Viet-Nam-pddvn/")

root = tk.Tk()
root.title("Web Scraper for PPDVN")

# Title Label
title_label = tk.Label(root, text="Web Scraper for PPDVN", font=("Helvetica", 16))
title_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

# Links
github_link = tk.Label(root, text="GitHub", font=("Tahoma", 12), fg="blue", cursor="hand2")
github_link.grid(row=1, column=0, columnspan=2)
github_link.bind("<Button-1>", lambda e: open_github())

# URL Input
tk.Label(root, text="Enter Title:").grid(row=2, column=0, padx=10, pady=10)
url_entry = tk.Entry(root, width=50)
url_entry.grid(row=2, column=1, padx=10, pady=10)

# Scrape Button
url_entry.bind("<Return>", lambda event: start_scraping_thread(url_entry.get()))
scrape_button = tk.Button(root, text="Start Scraping", command=lambda: start_scraping_thread(url_entry.get()))
scrape_button.grid(row=3, column=0, columnspan=2, pady=10)

# Run the application
root.mainloop()
