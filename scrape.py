from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    TimeoutException, 
    ElementClickInterceptedException,
    StaleElementReferenceException,
    NoSuchElementException
)
import time
import os
import sys
import signal
from pathlib import Path
import pandas as pd
from constants import LS_URL, RS_URL


class ParliamentScraper:
    """Scraper for Parliament Questions and Answers"""
    
    # Common Locators (work for both LS and RS)
    SEARCH_INPUT_XPATH = "//*[@id='input-with-icon-textfield']"
    EXPAND_ROW_SELECTOR = "button[aria-label*='expand']"  
    PDF_VIEWER_SELECTOR = ".rpv-core__viewer"
    DOWNLOAD_PDF_SELECTOR = "button[data-testid='get-file__download-button'], button[aria-label='Download']"
    NEXT_PAGE_SELECTOR = "button[aria-label='Go to next page']"
    LAST_PAGE_SELECTOR = "button[aria-label='Go to last page']"
    
    # Lok Sabha Specific Locators
    LS_LOCATORS = {
        'FACET_SEARCH_TAB': "/html/body/div[1]/div/div[3]/main/div/div[1]/main/div/div[2]/div[1]/div/button[2]",
        'FILTER_TYPE': "//*[@id='type']",
        'FILTER_MEMBERS_OPTION': "/html/body/div[9]/div[3]/ul/li[@data-value='members']",
        'FILTER_OPERATOR': "//*[@id='filter']",
        'FILTER_NOT_CONTAINS': "/html/body/div[9]/div[3]/ul/li[4]",
        'FILTER_VALUE_INPUT': "/html/body/div[1]/div/div[3]/main/div/div[1]/main/div/div[2]/div[2]/div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div/div/input",
        'APPLY_FILTER_BUTTON': "/html/body/div[1]/div/div[3]/main/div/div[1]/main/div/div[2]/div[2]/div[2]/div[1]/div[2]/div/div[2]/div[2]/button",
        'ENTRY_COUNT': "/html/body/div[1]/div/div[3]/main/div/div[1]/main/div/div[2]/div[2]/div[1]/div[2]/div[2]/div[1]/p",
        'PAGE_COUNT_DROPDOWN': "//*[@id='rows-per-page']",
        'PAGE_COUNT_100': "/html/body/div[9]/div[3]/ul/li[6]",
        'DOWNLOAD_BUTTON': "//*[@id='basic-button']",
        'EXPORT_EXCEL': "/html/body/div[9]/div[3]/ul/li"
    }
    
    # Rajya Sabha Specific Locators
    RS_LOCATORS = {
        'FACET_SEARCH_TAB': "//button[contains(text(), 'Facet Search on Questions')]",
        'FILTER_TYPE': "//*[@id='type']",
        'FILTER_MEMBERS_OPTION': "//li[contains(text(), 'Members')]",
        'FILTER_OPERATOR': "//*[@id='filter']",
        'FILTER_NOT_CONTAINS': "//li[contains(text(), 'Not Contains')]",
        'FILTER_VALUE_INPUT': "//input[contains(@class, 'MuiAutocomplete-input') and @role='combobox']",
        'APPLY_FILTER_BUTTON': "//button[@aria-label='button to apply custom filter']",
        'ENTRY_COUNT': "//p[contains(text(), 'Showing')]",
        'PAGE_COUNT_DROPDOWN': "//*[@id='rows-per-page']",
        'PAGE_COUNT_100': "//li[text()='100']",
        'EXPAND_BUTTON': "//button[@aria-label='expand row']",
        'TEXT_PDF_TAB': "//button[@role='tab' and contains(text(), 'Text PDF')]",
        'OPEN_BUTTON': "//div[@id='main-content']//a[@target='_blank']", 
        'DOWNLOAD_BUTTON': "//*[@id='basic-button']",
        'EXPORT_EXCEL': "//li[contains(text(), 'Export to Excel')]"
    }
    
    def __init__(self, search_term, source_name, headless=False):
        self.search_term = search_term
        self.source_name = source_name
        self.headless = headless
        self.driver = None
        self.wait = None
        self.long_wait = None
        self.download_dir = self._setup_download_directory()
        self.scraping_completed = False  
        self.processing_done = False 
        
        self.locators = self.LS_LOCATORS if source_name == 'LS' else self.RS_LOCATORS
        
        # Register signal handlers for graceful interruption
        signal.signal(signal.SIGINT, self._signal_handler)
        signal.signal(signal.SIGTERM, self._signal_handler)
        
    def _setup_download_directory(self):
        search_term_dir = self.search_term.strip()
        invalid_start_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|', '.']
        
        if search_term_dir and search_term_dir[0] in invalid_start_chars:
            search_term_dir = '_' + search_term_dir
        
        # directory structure: results/search_term/source_name/
        download_dir = Path.cwd() / "results" / search_term_dir / self.source_name
        download_dir.mkdir(parents=True, exist_ok=True)
        
        return str(download_dir)
    
    def _signal_handler(self, signum, frame):
        print("\nInterruption detected! Processing downloaded files before exit...")
        
        # Process whatever has been downloaded so far
        if not self.processing_done:
            self._process_results()
        
        if self.driver:
            try:
                self.driver.quit()
                print("Browser closed")
            except:
                pass
        
        print("\nScraping interrupted by user")
        print(f"Downloaded files have been processed and saved to: {self.download_dir}")
        sys.exit(0)
    
    def _initialize_driver(self):
        options = webdriver.ChromeOptions()
       
        prefs = {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "plugins.always_open_pdf_externally": True 
        }
        options.add_experimental_option("prefs", prefs)
        
        # Browser options
        if self.headless:
            options.add_argument('--headless=new')
        
        # Set window size for consistent behavior across environments
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--start-maximized')
        
        # Disable automation flags
        options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
        options.add_experimental_option('useAutomationExtension', False)
        
        self.driver = webdriver.Chrome(options=options)
        self.driver.maximize_window()
        
        # Initialize waits
        self.wait = WebDriverWait(self.driver, 20)
        self.long_wait = WebDriverWait(self.driver, 120)
        
        print(f"WebDriver initialized for {self.source_name}")
        print(f"Downloads will be saved to: {self.download_dir}")
    
    def _safe_click(self, element, use_js=False):
        max_retries = 3
        
        for attempt in range(max_retries):
            try:
                self.driver.execute_script(
                    "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", 
                    element
                )
                time.sleep(0.5)

                if use_js:
                    self.driver.execute_script("arguments[0].click();", element)
                else:
                    element.click()
                
                return True
                
            except (ElementClickInterceptedException, StaleElementReferenceException) as e:
                if attempt < max_retries - 1:
                    print(f"Click failed (attempt {attempt + 1})")
                    use_js = True
                    time.sleep(0.5)
                else:
                    raise e
        
        return False
    
    def _wait_and_click(self, by, value, timeout=20, scroll=True):
        wait = WebDriverWait(self.driver, timeout)
        element = wait.until(EC.element_to_be_clickable((by, value)))
        
        if scroll:
            self.driver.execute_script(
                "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", 
                element
            )
            time.sleep(0.3)
        
        self._safe_click(element)
        return element
    
    def _wait_for_element(self, by, value, timeout=20):
        wait = WebDriverWait(self.driver, timeout)
        return wait.until(EC.presence_of_element_located((by, value)))
    
    def _navigate_and_search(self, url):
        print(f"Starting scraping process for {self.source_name}\n")
        
        self.driver.get(url)
        print(f"Navigated to {self.source_name} URL")
        
        self._wait_and_click(By.XPATH, self.locators['FACET_SEARCH_TAB'])
        print("Clicked on Facet Search Tab")
        
        # Enter search term
        search_input = self._wait_for_element(By.XPATH, self.SEARCH_INPUT_XPATH)
        search_input.clear()
        search_input.send_keys(self.search_term)
        search_input.send_keys(Keys.RETURN)
        print(f"Entered search term: '{self.search_term}'")
        
    def _apply_member_filter(self):
        # Click filter type dropdown
        self._wait_and_click(By.XPATH, self.locators['FILTER_TYPE'])
        print("Opened filter type dropdown")
        
        # Select 'members' option
        self._wait_and_click(By.XPATH, self.locators['FILTER_MEMBERS_OPTION'])
        print("Selected 'Members' filter")
        
        # Click filter operator dropdown
        self._wait_and_click(By.XPATH, self.locators['FILTER_OPERATOR'])
        print("Opened filter operator dropdown")
        
        # Select 'not contains' option
        self._wait_and_click(By.XPATH, self.locators['FILTER_NOT_CONTAINS'])
        print("Selected 'Not Contains' operator")
        
        time.sleep(1)
        
        # Enter filter value
        filter_input = self._wait_for_element(By.XPATH, self.locators['FILTER_VALUE_INPUT'])
        
        if self.source_name == 'RS':
            self._safe_click(filter_input)
            time.sleep(0.5)
        
        filter_input.clear()
        filter_input.send_keys(self.search_term)
        
        filter_input.send_keys(Keys.RETURN)
        
        if self.source_name == 'RS':
            time.sleep(1)
        
        print(f"Entered filter value: '{self.search_term}'")
        
        # Apply filter
        self._wait_and_click(By.XPATH, self.locators['APPLY_FILTER_BUTTON'])
        print("Applied filter")
        
        time.sleep(2)
        
        try:
            entry_count = self._wait_for_element(By.XPATH, self.locators['ENTRY_COUNT'], timeout=5)
            print(f"Filter applied successfully")
            print(f"\n{self.source_name} Entry Count: {entry_count.text}\n")
        except:
            print(f"Filter applied successfully")
            print(f"\n{self.source_name}: Results loaded (entry count not available)\n")
        
    def _set_max_entries_per_page(self):
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        
        # Click dropdown to open options
        dropdown = self._wait_for_element(By.XPATH, self.locators['PAGE_COUNT_DROPDOWN'])
        self._safe_click(dropdown)
        time.sleep(0.5)
        
        # Click 100 option
        self._wait_and_click(By.XPATH, self.locators['PAGE_COUNT_100'])
        print("Set entries per page to 100")
        time.sleep(2) 
        
    def _download_excel_for_page(self):
        body = self.driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.HOME)
        time.sleep(0.5)
        
        download_btn = self._wait_for_element(By.XPATH, self.locators['DOWNLOAD_BUTTON'])
        self._safe_click(download_btn, use_js=True)
        
        self._wait_and_click(By.XPATH, self.locators['EXPORT_EXCEL'])
        print("Excel file download initiated")
        
        time.sleep(2)
        
    def _download_pdfs_for_page(self, page_num):
        arrows = self.wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, self.EXPAND_ROW_SELECTOR))
        )
        
        total_rows = len(arrows)
        print(f"\nPage {page_num}: Found {total_rows} rows")
        
        first_pdf_on_page = True
        
        for i, arrow in enumerate(arrows, 1):
            try:
                self.driver.execute_script(
                    "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", 
                    arrow
                )
                time.sleep(0.3)
                
                self._safe_click(arrow, use_js=True)
                print(f"[{i}/{total_rows}] Expanded row")
                
                if self.source_name == 'RS':
                    # RS
                    time.sleep(1.5) 
                    
                    try:
                        text_pdf_tab = self.wait.until(
                            EC.element_to_be_clickable((By.XPATH, self.locators['TEXT_PDF_TAB']))
                        )
                        self._safe_click(text_pdf_tab, use_js=True)
                        time.sleep(2)
                    except:
                        pass
                    
                    # Scroll down to find the PDF link
                    self.driver.execute_script("window.scrollBy(0, 400);")
                    time.sleep(1.5)
                    
                    # Find PDF URL from iframes or main document
                    iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                    pdf_url = None
                    
                    for iframe in iframes:
                        try:
                            self.driver.switch_to.frame(iframe)
                            all_links = self.driver.find_elements(By.TAG_NAME, "a")
                            for link in all_links:
                                href = link.get_attribute('href')
                                if href and '.pdf' in href:
                                    pdf_url = href
                                    break
                            self.driver.switch_to.default_content()
                            if pdf_url:
                                break
                        except:
                            self.driver.switch_to.default_content()
                            continue
                    
                    if not pdf_url:
                        all_links = self.driver.find_elements(By.TAG_NAME, "a")
                        for link in all_links:
                            href = link.get_attribute('href')
                            if href and '.pdf' in href:
                                pdf_url = href
                                break
                    
                    if not pdf_url:
                        raise Exception("Could not find PDF URL")
                    
                    self.driver.execute_script(f"window.open('{pdf_url}', '_blank');")
                    time.sleep(1.5)
                    
                    if len(self.driver.window_handles) > 1:
                        self.driver.switch_to.window(self.driver.window_handles[-1])
                        self.driver.close()
                        self.driver.switch_to.window(self.driver.window_handles[0])
                    
                    print(f"[{i}/{total_rows}] PDF downloaded")
                    
                else:
                    # LS
                    self.long_wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, self.PDF_VIEWER_SELECTOR))
                    )
                    
                    if first_pdf_on_page:
                        time.sleep(3)
                        first_pdf_on_page = False
                    else:
                        time.sleep(0.5)
                    
                    download_btn = self.long_wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, self.DOWNLOAD_PDF_SELECTOR))
                    )
                    self._safe_click(download_btn, use_js=True)
                    print(f"[{i}/{total_rows}] PDF downloaded")
                
                time.sleep(1)

                self._safe_click(arrow, use_js=True)
                time.sleep(0.3)
                
            except Exception as e:
                print(f"[{i}/{total_rows}] Error: {str(e)[:100]}")
                try:
                    if arrow.get_attribute('aria-expanded') == 'true':
                        self._safe_click(arrow, use_js=True)
                except:
                    pass
                continue
    
    def _navigate_pages(self):
        page_num = 1
        
        while True:
            print(f"\nProcessing Page {page_num}\n")
            
            # Download Excel for this page
            self._download_excel_for_page()
            
            # Download PDFs for this page
            self._download_pdfs_for_page(page_num)
            
            # Check if we can go to next page
            try:
                next_btn = self.driver.find_element(By.CSS_SELECTOR, self.NEXT_PAGE_SELECTOR)
                last_btn = self.driver.find_element(By.CSS_SELECTOR, self.LAST_PAGE_SELECTOR)
                
                # If last page button is disabled, already on last page
                if last_btn.get_attribute("disabled"):
                    print(f"\nReached last page (Page {page_num})")
                    break
                
                # If next button is enabled, click it
                if not next_btn.get_attribute("disabled"):
                    self._safe_click(next_btn, use_js=True)
                    print(f"\n→ Navigating to Page {page_num + 1}")
                    time.sleep(2) 
                    page_num += 1
                else:
                    print(f"\nReached last page (Page {page_num})")
                    break
                    
            except NoSuchElementException:
                print(f"\nNo more pages available")
                break
    
    def _process_results(self):
        if self.processing_done:
            print("Results already processed")
            return
        
        print("Processing Results")
        
        results_dir = Path(self.download_dir)
        xls_files = list(results_dir.glob('*.xls'))
        
        if not xls_files:
            print("No .xls files found to process")
            self.processing_done = True
            return
        
        print(f"Found {len(xls_files)} Excel files to combine")
        
        all_dataframes = []
        
        for xls_file in xls_files:
            try:
                tables = pd.read_html(str(xls_file))
                if tables:
                    df = tables[0]
                    all_dataframes.append(df)
                    print(f"Processed: {xls_file.name} ({len(df)} rows)")
            except Exception as e:
                print(f"Error processing {xls_file.name}: {e}")
        
        if all_dataframes:
            combined_df = pd.concat(all_dataframes, ignore_index=True)
            
            original_count = len(combined_df)
            combined_df = combined_df.drop_duplicates()
            duplicates_removed = original_count - len(combined_df)
            
            # Save as Excel
            output_file = results_dir / f"{self.search_term}_{self.source_name}_combined.xlsx"
            combined_df.to_excel(output_file, index=False, engine='openpyxl')
            
            print(f"\nCombined {len(all_dataframes)} files")
            print(f"Total rows: {len(combined_df)}")
            if duplicates_removed > 0:
                print(f"Removed {duplicates_removed} duplicate rows")
            print(f"Output: {output_file.name}")
            
            # Cleanup original .xls files
            print("Cleaning up temporary files")
            for xls_file in xls_files:
                try:
                    xls_file.unlink()
                    print(f"Deleted: {xls_file.name}")
                except Exception as e:
                    print(f"Error deleting {xls_file.name}: {e}")
            
            print(f"\nCleaned up {len(xls_files)} temporary files")
        else:
            print("No data to combine")
        
        self.processing_done = True
    
    def scrape(self):
        try:
            self._initialize_driver()
            
            url = LS_URL if self.source_name == 'LS' else RS_URL
            
            self._navigate_and_search(url)
            self._apply_member_filter()
            self._set_max_entries_per_page()
            self._navigate_pages()
            
            time.sleep(10)
   
            self.scraping_completed = True
            self._process_results()
            print(f"\n{self.source_name} Scraping Completed Successfully!")
            
        except KeyboardInterrupt:
            print("\nKeyboard interrupt detected")
            raise
            
        except Exception as e:
            print(f"\nError during scraping: {e}")
            print(f"\nProcessing downloaded files before exit")
            
            try:
                if not self.processing_done:
                    self._process_results()
                print(f"\nDownloaded files have been processed despite the error")
                print(f"Check results at: {self.download_dir}")
            except Exception as process_error:
                print(f"Error during processing: {process_error}")
            
            import traceback
            traceback.print_exc()
            
        finally:
            if self.driver:
                try:
                    self.driver.quit()
                    print(f"\nBrowser closed")
                except Exception as e:
                    print(f"Error closing browser: {e}")
            
            if not self.scraping_completed and not self.processing_done:
                print("Scraping did not complete normally")
                print("Processing partial downloads...")
                try:
                    self._process_results()
                    print(f"\nPartial results saved to: {self.download_dir}")
                except Exception as e:
                    print(f"Could not process partial results: {e}")

def main():
    print("Parliament Questions & Answers Scraper")
    
    search_term = input("Enter search term: ").strip()
    
    if not search_term:
        print("Search term cannot be empty!")
        return
    
    print("\nSelect source:")
    print("1. LS (Lok Sabha)")
    print("2. RS (Rajya Sabha)")
    print("3. Both")
    
    source_choice = input("\nEnter choice (1/2/3): ").strip()
    
    if source_choice not in ['1', '2', '3']:
        print("Invalid choice!")
        return
    
    if source_choice in ['1', '3']:
        scraper = ParliamentScraper(search_term, "LS", headless=False)
        scraper.scrape()
    
    if source_choice in ['2', '3']:
        scraper = ParliamentScraper(search_term, "RS", headless=False)
        scraper.scrape()
    
    print("Scraping Completed!")

if __name__ == "__main__":
    main()
