import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import re  # Regular expressions for cleaning text


def get_university_data(driver, url):
    """
    Navigates to a single university page and extracts key data points.
    Returns a dictionary with the extracted information.
    """
    print(f"  -> Scraping: {url}")
    driver.get(url)
    data = {'URL': url}

    try:
        # Wait for the main stats container to be visible
        stats_container = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.stats-wrapper"))
        )

        # --- Helper function to extract stats ---
        def extract_stat(label):
            try:
                # Find the div containing the label text, then get the value from the sibling div
                # This XPath is more specific to avoid errors.
                value_element = stats_container.find_element(
                    By.XPATH,
                    f".//div[contains(@class, '_label') and contains(text(), '{label}')]/following-sibling::div[contains(@class, '_value')]"
                )
                # Clean the text to get only numbers
                value = re.sub(r'[^\d]', '', value_element.text)
                return int(value) if value else 'Not Found'
            except Exception:
                return 'Not Found'

        # Extracting data based on your image
        data['Total Students'] = extract_stat('Total students')
        data['International Students'] = extract_stat('International students')
        data['Total Faculty Staff'] = extract_stat('Total faculty staff')

    except Exception as e:
        print(f"    - Could not load stats container for {url}.")

    return data


# --- Main Script ---
if __name__ == "__main__":
    # Use the main URL, which redirects to the latest ranking
    BASE_URL = "https://www.topuniversities.com/world-university-rankings"

    print("Setting up the browser driver...")
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless') # Keep this commented out for testing to see the browser
    options.add_argument('--log-level=3')
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()  # Maximize window to ensure all elements are visible

    print(f"Navigating to the main rankings page: {BASE_URL}")
    driver.get(BASE_URL)

    # --- NEW & CRITICAL STEP: Handle the cookie consent banner ---
    try:
        print("Looking for the cookie consent banner...")
        accept_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))
        )
        print("Cookie banner found. Clicking 'Accept All Cookies'.")
        accept_button.click()
        time.sleep(2)  # Give a moment for the banner to disappear
    except Exception as e:
        print("Cookie banner not found or could not be clicked, continuing anyway.")

    # --- Step 1: Get all university links from the main page ---
    university_links = []
    print("Finding all university links on the page...")

    try:
        # Wait for the first university row to be loaded before we start scrolling
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.ind-row"))
        )

        # Scroll down to load all universities on the page
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            print("  -> Scrolling to load more universities...")
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(4)  # Increased wait time for content to load properly
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        print("All universities loaded. Extracting links...")
        uni_elements = driver.find_elements(By.CSS_SELECTOR, "a.uni-link")
        for element in uni_elements:
            university_links.append(element.get_attribute('href'))

        # Optional: For a quick test, uncomment the line below to only scrape 10 universities
        # university_links = university_links[:10]

        print(f"Successfully found {len(university_links)} university links.")

    except Exception as e:
        print(f"An error occurred while trying to find university links: {e}")
        driver.quit()
        exit()

    # --- Step 2: Loop through each link and scrape data ---
    all_university_data = []
    total_universities = len(university_links)

    for i, link in enumerate(university_links):
        print(f"\nProcessing University {i + 1} of {total_universities}...")
        uni_data = get_university_data(driver, link)

        try:
            title = driver.title
            uni_data['University Name'] = title.split('|')[0].strip()
        except:
            uni_data['University Name'] = "Name Not Found"

        all_university_data.append(uni_data)
        time.sleep(1)

        # --- Step 3: Save the data to an Excel file ---
    print("\n✅ Scraping complete! Saving data to Excel...")
    df = pd.DataFrame(all_university_data)

    column_order = ['University Name', 'Total Students', 'International Students', 'Total Faculty Staff', 'URL']
    # Filter out columns that may not exist in the dataframe before reordering
    existing_columns = [col for col in column_order if col in df.columns]
    df = df[existing_columns]

    output_filename = 'university_data.xlsx'
    df.to_excel(output_filename, index=False)

    print(f"Data successfully saved to '{output_filename}'")

    driver.quit()