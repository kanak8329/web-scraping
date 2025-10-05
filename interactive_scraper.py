import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import re


# --- BACKEND SCRAPING LOGIC ---

class UniversityScraper:
    def __init__(self, status_callback):
        self.BASE_URL = "https://www.topuniversities.com/world-university-rankings"
        self.update_status = status_callback
        self.driver = None

    def _setup_driver(self):
        self.update_status("Setting up browser driver...")
        service = Service(ChromeDriverManager().install())
        options = webdriver.ChromeOptions()
        options.add_argument("start-maximized")
        options.add_argument('--log-level=3')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        self.driver = webdriver.Chrome(service=service, options=options)
        self.update_status("Driver setup complete. Browser window will open.")

    def _handle_popups(self):
        self.update_status("Actively checking for pop-ups for 15 seconds...")
        start_time = time.time()
        while time.time() - start_time < 15:
            try:
                cookie_button = self.driver.find_element(By.ID, "onetrust-accept-btn-handler")
                cookie_button.click()
                self.update_status("  > Cookie banner accepted.")
                time.sleep(1)
                continue
            except:
                pass

            try:
                close_button = self.driver.find_element(By.CSS_SELECTOR, "div.qs-modal-close")
                close_button.click()
                self.update_status("  > Promotional pop-up closed.")
                time.sleep(1)
                continue
            except:
                pass

            time.sleep(0.5)
        self.update_status("Pop-up check complete.")

    def _get_university_links(self):
        self.update_status(f"Navigating to {self.BASE_URL}")
        self.driver.get(self.BASE_URL)

        self._handle_popups()

        self.update_status("Waiting for university list to load...")
        WebDriverWait(self.driver, 45).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.ind-row"))
        )
        self.update_status("Main content loaded successfully.")

        last_height = self.driver.execute_script("return document.body.scrollHeight")
        while True:
            self.update_status("Scrolling down to load all universities...")
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(4)
            new_height = self.driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        uni_elements = self.driver.find_elements(By.CSS_SELECTOR, "a.uni-link")
        links = [elem.get_attribute('href') for elem in uni_elements]
        self.update_status(f"Found {len(links)} university links.")
        return links

    def _extract_details(self, url):
        self.driver.get(url)
        data = {'URL': url, 'University Name': 'Not Found'}
        try:
            data['University Name'] = self.driver.title.split('|')[0].strip()
        except:
            pass

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
            self.update_status(f"  - No detailed stats found for {data['University Name']}")
        return data

    def run(self):
        try:
            self._setup_driver()
            links = self._get_university_links()

            all_data = []
            total = len(links)
            for i, link in enumerate(links):
                self.update_status(f"Processing ({i + 1}/{total}): {link.split('/')[-1]}")
                details = self._extract_details(link)
                all_data.append(details)
                time.sleep(0.5)

            self.update_status("Scraping complete. Creating Excel file...")
            df = pd.DataFrame(all_data)
            cols_order = ['University Name', 'Total Students (Total)', 'Total Students (UG students)',
                          'Total Students (PG students)', 'International Students (Total)',
                          'International Students (UG students)', 'International Students (PG students)',
                          'Total Faculty Staff (Total)', 'Total Faculty Staff (Domestic staff)',
                          'Total Faculty Staff (Int\'l staff)', 'URL']
            df = df.reindex(columns=cols_order)
            filename = 'university_student_staff_data.xlsx'
            df.to_excel(filename, index=False)
            self.update_status(f"\nSUCCESS! ✅\nData saved to '{filename}'")

        except Exception as e:
            self.update_status(f"\nERROR: An unexpected error occurred.\nDetails: {e}")
        finally:
            if self.driver:
                self.driver.quit()
                self.update_status("Browser closed.")


# --- GUI APPLICATION ---

class ScraperApp:
    def __init__(self, root):
        self.root = root
        root.title("University Data Scraper")
        root.geometry("700x500")
        self.style = ttk.Style()
        self.style.configure("TButton", font=("Helvetica", 12, "bold"), padding=10)
        self.style.configure("TLabel", font=("Helvetica", 14, "bold"))
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        title_label = ttk.Label(main_frame, text="TopUniversities.com Scraper")
        title_label.pack(pady=10)
        self.log_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=20, width=80)
        self.log_text.pack(pady=5, fill=tk.BOTH, expand=True)
        self.log_text.insert(tk.END,
                             "Click 'Start Scraping' to begin.\nA Chrome window will open to perform the task.\n\n")
        self.start_button = ttk.Button(main_frame, text="Start Scraping", command=self.start_scraping_thread)
        self.start_button.pack(pady=10, fill=tk.X)

    def update_status(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def start_scraping_thread(self):
        self.start_button.config(state=tk.DISABLED, text="Scraping in Progress...")
        self.log_text.delete('1.0', tk.END)
        self.update_status("Starting the scraping process...")

        # This is the corrected line
        scraper_thread = threading.Thread(target=self.run_scraper, daemon=True)
        scraper_thread.start()

    def run_scraper(self):
        scraper = UniversityScraper(self.update_status)
        scraper.run()
        self.start_button.config(state=tk.NORMAL, text="Start Scraping")


# --- Run the Application ---
if __name__ == "__main__":
    root = tk.Tk()
    app = ScraperApp(root)
    root.mainloop()