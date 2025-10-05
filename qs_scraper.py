"""
QS full scraper: visits every university profile from
https://www.topuniversities.com/world-university-rankings and extracts:
- Total Students
- UG Students
- PG Students
- International Students
- Total faculty staff
- Domestic staff

Saves incremental CSV and final Excel.
"""

import time
import csv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import (
    NoSuchElementException, TimeoutException, StaleElementReferenceException,
    WebDriverException
)
from tqdm import tqdm
import traceback

# ---------- CONFIG ----------
START_INDEX = 0            # set >0 to resume from a particular row count of the link list (0 = start)
INCREMENTAL_CSV = "qs_rankings_progress.csv"
FINAL_XLSX = "qs_rankings_full.xlsx"
BASE_URL = "https://www.topuniversities.com/world-university-rankings"
HEADLESS = False           # set True to run headless (useful on servers)
MAX_RETRIES = 3
DELAY_BETWEEN_OPEN = 2.5   # seconds between opening profile pages (politeness)
LOAD_MORE_WAIT = 2         # seconds after clicking Load More
# ----------------------------

def setup_driver():
    options = Options()
    if HEADLESS:
        options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    # reduce detection
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def click_cookie_if_present(driver, wait):
    try:
        cookie_btn = wait.until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler")), timeout=5)
        cookie_btn.click()
        time.sleep(1)
    except Exception:
        # no cookie popup or timed out
        pass

def load_all_rows(driver, wait, max_click_failures=5):
    """Click 'Load more' until it disappears or cannot be clicked."""
    failures = 0
    while True:
        try:
            # attempt to find "Load more" button (class used on site)
            load_more = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.js-rankings-load-more")), timeout=10)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", load_more)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", load_more)
            time.sleep(LOAD_MORE_WAIT)
            failures = 0
        except TimeoutException:
            # no button found — assume all loaded
            break
        except Exception as e:
            failures += 1
            print(f"Load more click failed ({failures}) — {e}")
            time.sleep(2)
            if failures >= max_click_failures:
                print("Too many failures clicking Load More; stopping attempts.")
                break

def collect_university_links(driver):
    """Return list of (rank_text, uni_name, profile_url) for all rows visible."""
    rows = driver.find_elements(By.CSS_SELECTOR, "div.rankings-table__row")
    links = []
    for r in rows:
        try:
            rank = r.find_element(By.CSS_SELECTOR, ".rankings-table__rank").text.strip()
        except Exception:
            rank = ""
        try:
            name = r.find_element(By.CSS_SELECTOR, ".rankings-table__title").text.strip()
        except Exception:
            name = ""
        try:
            a = r.find_element(By.TAG_NAME, "a")
            href = a.get_attribute("href")
        except Exception:
            href = ""
        if href:
            links.append((rank, name, href))
    return links

def extract_stats_from_profile(driver, wait):
    """
    Find uni-stats-card elements and map titles to numbers.
    Use robust fallback: search for title text in page as well.
    Returns dict of target fields (some may be empty strings).
    """
    desired = {
        "Total students": "",
        "UG students": "",
        "PG students": "",
        "International students": "",
        "Total faculty staff": "",
        "Domestic staff": ""
    }

    try:
        # Wait briefly for stats cards to appear; they may be inside X blocks
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".uni-stats-card")), timeout=8)
        cards = driver.find_elements(By.CSS_SELECTOR, ".uni-stats-card")
        for c in cards:
            try:
                title = c.find_element(By.CSS_SELECTOR, ".uni-stats-card__title").text.strip()
                value = c.find_element(By.CSS_SELECTOR, ".uni-stats-card__number").text.strip()
                if not title:
                    continue
                # Normalize title and map
                t = title.lower()
                if "total students" in t:
                    desired["Total students"] = value
                elif "ug students" in t or "undergraduate students" in t:
                    desired["UG students"] = value
                elif "pg students" in t or "postgraduate students" in t:
                    desired["PG students"] = value
                elif "international students" in t:
                    desired["International students"] = value
                elif "total faculty staff" in t or "faculty staff" in t:
                    desired["Total faculty staff"] = value
                elif "domestic staff" in t:
                    desired["Domestic staff"] = value
            except Exception:
                continue
    except Exception:
        # fallback: attempt to search for labels in page text
        page_text = driver.page_source.lower()
        for key in list(desired.keys()):
            if key in page_text:
                # attempt to extract nearby number using simple heuristics (xpath near text)
                try:
                    # find element that contains the key text
                    el = driver.find_element(By.XPATH, f"//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{key}')]")
                    # try to find following sibling number
                    try:
                        num = el.find_element(By.XPATH, ".//following::div[1]").text.strip()
                    except Exception:
                        num = ""
                    if num:
                        desired[key] = num
                except Exception:
                    pass

    return desired

def append_to_csv(row, csv_path=INCREMENTAL_CSV):
    header = ["Rank", "University", "Profile URL",
              "Total Students", "UG Students", "PG Students", "International Students",
              "Total Faculty Staff", "Domestic Staff"]
    file_exists = False
    try:
        with open(csv_path, "r", encoding="utf-8") as fr:
            file_exists = True
    except FileNotFoundError:
        file_exists = False

    with open(csv_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(header)
        writer.writerow(row)

def main():
    driver = setup_driver()
    wait = WebDriverWait(driver, 15)
    driver.get(BASE_URL)
    time.sleep(3)

    click_cookie_if_present(driver, wait)
    time.sleep(1)

    print("Loading all rows (this may take a while)...")
    load_all_rows(driver, wait)

    # Collect all links
    links = collect_university_links(driver)
    print(f"Collected {len(links)} university links.")

    # If you are resuming, skip to START_INDEX
    total_links = len(links)
    if START_INDEX >= total_links:
        print("START_INDEX >= total links. Nothing to do.")
        driver.quit()
        return

    # Iterate through links and extract
    for idx in tqdm(range(START_INDEX, total_links), desc="Universities"):
        rank_text, uni_name, url = links[idx]
        # Retry logic for each profile
        success = False
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                # open profile in a new tab
                driver.execute_script("window.open(arguments[0]);", url)
                driver.switch_to.window(driver.window_handles[-1])
                time.sleep(DELAY_BETWEEN_OPEN)  # politeness

                stats = extract_stats_from_profile(driver, wait)

                row = [
                    rank_text,
                    uni_name,
                    url,
                    stats.get("Total students", ""),
                    stats.get("UG students", ""),
                    stats.get("PG students", ""),
                    stats.get("International students", ""),
                    stats.get("Total faculty staff", ""),
                    stats.get("Domestic staff", "")
                ]
                append_to_csv(row)
                success = True

                # close tab and return
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

                # short sleep to avoid hammering
                time.sleep(1)
                break
            except (WebDriverException, Exception) as e:
                print(f"Attempt {attempt} failed for {uni_name}: {e}")
                traceback.print_exc()
                # close any extra tabs if open
                try:
                    while len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                except Exception:
                    pass
                time.sleep(2 * attempt)
        if not success:
            # write an empty row indicating failure
            print(f"Failed to scrape {uni_name} after {MAX_RETRIES} attempts. Writing blank row.")
            append_to_csv([rank_text, uni_name, url, "", "", "", "", "", ""])

    # finished: create final Excel
    try:
        df = pd.read_csv(INCREMENTAL_CSV)
        df.to_excel(FINAL_XLSX, index=False)
        print(f"Saved final Excel to {FINAL_XLSX}")
    except Exception as e:
        print("Could not write final Excel:", e)

    driver.quit()

if __name__ == "__main__":
    main()
