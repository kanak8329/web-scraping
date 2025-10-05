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
    The final and most robust version of the scraper, designed to handle iframes.
    """

    def __init__(self):
        self.BASE_URL = "https://www.topuniversities.com/world-university-rankings"
        self.driver = None

    def _setup_driver(self):
        print("➡️ Setting up stealth browser driver...")
        service = Service(ChromeDriverManager().install())
        options = webdriver.ChromeOptions()
        options.add_argument("start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        self.driver = webdriver.Chrome(service=service, options=options)
        stealth(self.driver, languages=["en-US", "en"], vendor="Google Inc.", platform="Win32",
                webgl_vendor="Intel Inc.", renderer="Intel Iris OpenGL Engine", fix_hairline=True)
        print("✅ Stealth driver setup complete.")

    # --- COMPLETELY REWRITTEN FUNCTION ---
    def _handle_popups(self):
        """
        Handles pop-ups by specifically looking for and switching into the cookie banner's iframe.
        """
        print("🔎 Looking for the cookie banner iframe...")
        try:
            # Wait for the iframe itself to be present on the page
            iframe_wait = WebDriverWait(self.driver, 20)
            cookie_iframe = iframe_wait.until(
                EC.presence_of_element_located((By.XPATH, "//iframe[contains(@title, 'Modal Dialog')]"))
            )

            # Switch the driver's focus into the iframe
            self.driver.switch_to.frame(cookie_iframe)
            print("   - Switched into iframe. Looking for accept button...")

            # Now, inside the iframe, look for the accept button
            button_wait = WebDriverWait(self.driver, 10)
            accept_button = button_wait.until(
                EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))
            )

            # Click the button
            accept_button.click()
            print("   - Cookie banner accepted.")

        except Exception as e:
            print(f"   - Could not find or interact with the cookie iframe. The website may have changed. Error: {e}")

        finally:
            # CRITICAL: Always switch back to the main page content
            self.driver.switch_to.default_content()
            print("   - Switched back to main page content.")
            print("✅ Pop-up check complete.")

    def _get_university_links(self):
        print(f"🌍 Navigating to the main rankings page...")
        self.driver.get(self.BASE_URL)
        self._handle_popups()

        print("⏳ Waiting for the main university list to load...")
        WebDriverWait(self.driver, 45).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.ind-row")))
        print("✅ Main content loaded.")

        last_height = self.driver.execute_script("return document.body.scrollHeight")
        while True:
            print("📜 Scrolling down to load all universities...")
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(4)
            new_height = self.driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height: break
            last_height = new_height

        links = [elem.get_attribute('href') for elem in self.driver.find_elements(By.CSS_SELECTOR, "a.uni-link")]
        print(f"👍 Found {len(links)} university links.")
        return links

    def _extract_page_details(self, url):
        self.driver.get(url)
        data = {'URL': url, 'University Name': self.driver.title.split('|')[0].strip()}
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
                        name, value = ratio.find_element(By.CSS_SELECTOR, "div._name").text, ratio.find_element(
                            By.CSS_SELECTOR, "div._value").text
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
        try:
            self._setup_driver()
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
            cols_order = ['University Name', 'Total Students (Total)', 'Total Students (UG students)',
                          'Total Students (PG students)', 'International Students (Total)',
                          'International Students (UG students)', 'International Students (PG students)',
                          'Total Faculty Staff (Total)', 'Total Faculty Staff (Domestic staff)',
                          'Total Faculty Staff (Int\'l staff)', 'URL']
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