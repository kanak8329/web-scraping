import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
from selenium_stealth import stealth


class UniversityScraper:
    """
    A slow, patient, and reliable scraper designed to mimic a human user
    and defeat anti-scraping measures by being deliberate.
    """

    def __init__(self):
        self.BASE_URL = "https://www.topuniversities.com/world-university-rankings"
        self.driver = None

    def _setup_driver(self):
        """Initializes a stealthy, visible Chrome browser."""
        print("➡️ Setting up the browser...")
        service = Service(ChromeDriverManager().install())
        options = webdriver.ChromeOptions()
        options.add_argument("start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        self.driver = webdriver.Chrome(service=service, options=options)

        # Apply stealth settings to make the browser look like a real person's
        stealth(self.driver,
                languages=["en-US", "en"],
                vendor="Google Inc.",
                platform="Win32",
                webgl_vendor="Intel Inc.",
                renderer="Intel Iris OpenGL Engine",
                fix_hairline=True,
                )

        print("✅ Browser is ready.")

    def _handle_initial_page_load(self):
        """Navigates and handles the cookie pop-up with extreme patience."""
        print(f"🌍 Navigating to the main rankings page...")
        self.driver.get(self.BASE_URL)

        print("⏳ Patiently waiting for the cookie pop-up (up to 40 seconds)...")
        try:
            # Use a very long wait time to ensure the pop-up has time to appear
            wait = WebDriverWait(self.driver, 40)
            cookie_button = wait.until(
                EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))
            )

            print("   - Pop-up found. Clicking 'Accept'.")
            # Use a JavaScript click, which is more reliable
            self.driver.execute_script("arguments[0].click();", cookie_button)
            # Pause deliberately after clicking
            time.sleep(3)
            print("✅ Pop-up handled.")

        except Exception:
            print("   - Cookie pop-up did not appear or was already handled. Continuing.")

    def _get_university_links(self):
        """Scrolls down the page very slowly to ensure all universities load."""
        print("⏳ Waiting for the main university list to appear...")
        WebDriverWait(self.driver, 40).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.ind-row"))
        )
        print("✅ Main list appeared. Starting to scroll.")

        last_height = self.driver.execute_script("return document.body.scrollHeight")
        while True:
            print("📜 Scrolling down slowly...")
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            # This long pause is the key to reliability
            time.sleep(5)
            new_height = self.driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        uni_elements = self.driver.find_elements(By.CSS_SELECTOR, "a.uni-link")
        links = [elem.get_attribute('href') for elem in uni_elements]
        print(f"👍 Found {len(links)} university links.")
        return links

    def _extract_page_details(self, url):
        """Visits a single university page and extracts the required data."""
        self.driver.get(url)
        data = {'URL': url}
        try:
            data['University Name'] = self.driver.title.split('|')[0].strip()
        except:
            data['University Name'] = "Name Not Found"

        try:
            stats_container = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.stats-wrapper")))

            def get_stat_box(label_text):
                try:
                    return stats_container.find_element(By.XPATH,
                                                        f".//div[contains(@class, '_label') and contains(text(), '{label_text}')]/ancestor::div[contains(@class, '_stat-box')]")
                except:
                    return None

            def process_box(box, data_prefix):
                if not box: return
                try:
                    data[f'{data_prefix} (Total)'] = re.sub(r'[^\d]', '',
                                                            box.find_element(By.CSS_SELECTOR, "div._value").text)
                except:
                    pass
                try:
                    for ratio in box.find_elements(By.CSS_SELECTOR, "div._ratio-box"):
                        name = ratio.find_element(By.CSS_SELECTOR, "div._name").text
                        value = ratio.find_element(By.CSS_SELECTOR, "div._value").text
                        data[f'{data_prefix} ({name})'] = value
                except:
                    pass

            process_box(get_stat_box('Total students'), 'Total Students')
            process_box(get_stat_box('International students'), 'International Students')
            process_box(get_stat_box('Total faculty staff'), 'Total Faculty Staff')
        except:
            pass
        return data

    def scrape(self):
        """The main method that orchestrates the entire scraping process."""
        try:
            self._setup_driver()
            self._handle_initial_page_load()
            links = self._get_university_links()

            all_data = []
            total_links = len(links)
            for i, link in enumerate(links):
                print(f"⚙️ Processing ({i + 1}/{total_links}): {link.split('/')[-1]}")
                details = self._extract_page_details(link)
                all_data.append(details)

            print("\n" + "=" * 50)
            print("💾 Scraping complete. Saving data to Excel file...")
            df = pd.DataFrame(all_data)
            cols_order = [
                'University Name', 'Total Students (Total)', 'Total Students (UG students)',
                'Total Students (PG students)',
                'International Students (Total)', 'International Students (UG students)',
                'International Students (PG students)',
                'Total Faculty Staff (Total)', 'Total Faculty Staff (Domestic staff)',
                'Total Faculty Staff (Int\'l staff)', 'URL'
            ]
            df = df.reindex(columns=cols_order)
            filename = 'university_data.xlsx'
            df.to_excel(filename, index=False)
            print(f"🎉 SUCCESS! Data for {len(all_data)} universities saved to '{filename}'")

        except Exception as e:
            print(f"\n❌ ERROR: An unexpected error occurred.\nDetails: {e}")
        finally:
            if self.driver:
                self.driver.quit()
                print("🔒 Browser closed.")


if __name__ == "__main__":
    scraper = UniversityScraper()
    scraper.scrape()