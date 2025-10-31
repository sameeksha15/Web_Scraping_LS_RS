from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import os
from constants import LS_URL, RS_URL
import pandas as pd

def search_and_scrape(driver, url, search_term, source_name):
    """Navigate to URL, enter search term, and scrape results."""

    try:
        print(f"\nScraping {source_name}...")
        driver.get(url)
        
        # navigate to facet search tab using xpath
        facet_search_tab = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[3]/main/div/div[1]/main/div/div[2]/div[1]/div/button[2]"))
        )
        facet_search_tab.click()
        print("[1] Clicked on Facet Search Tab")

        # Wait for page to load and find the search input field using xpath
        search_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='input-with-icon-textfield']"))
        )
        print("[2] Found Search Input Field")
        # Enter search term
        search_input.clear()
        search_input.send_keys(search_term)
        search_input.send_keys(Keys.RETURN)
        print(f"[3] Entered Search Term: {search_term}")

        # Add delay before scraping table
        time.sleep(5)

        # add filters using xpath
        add_filter_button = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='type']"))
        )
        add_filter_button.click()
        print("[4] Clicked on Add Filter Button")
        time.sleep(1)
        # dropdown to select filter option from li where data-value = 'members'
        filter_option = WebDriverWait(driver, 20).until(    
            EC.presence_of_element_located((By.XPATH, "/html/body/div[9]/div[3]/ul/li[@data-value='members']"))
        )
        filter_option.click()
        print("[5] Selected Filter Option: Members")

        time.sleep(2)

        # click on filter operator dropdown
        filter_operator = WebDriverWait(driver, 20).until(    
            EC.presence_of_element_located((By.XPATH, "//*[@id='filter']"))
        )
        filter_operator.click()
        time.sleep(1)
        # select not contains from dropdown (data-value='4')
        not_contains_option = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[9]/div[3]/ul/li[4]"))
        )
        not_contains_option.click()
        print("[6] Selected Filter Operator: Not Contains")

        time.sleep(3)


        # enter search term in filter value input by XPath
        filter_value_input = WebDriverWait(driver, 20).until(    
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[3]/main/div/div[1]/main/div/div[2]/div[2]/div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div/div/input"))
        )

        filter_value_input.click()
        filter_value_input.send_keys(search_term)
        filter_value_input.send_keys(Keys.RETURN)
        print(f"[7] Entered Filter Value: {search_term}")

        time.sleep(3)

        #click on apply filter button
        apply_filter_button = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[3]/main/div/div[1]/main/div/div[2]/div[2]/div[2]/div[1]/div[2]/div/div[2]/div[2]/button"))
        )
        apply_filter_button.click()
        print("[8] Clicked on Apply Filter Button")

        time.sleep(5)  # wait for filter options to load

        # read this p tag using xpath
        entry_count = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[3]/main/div/div[1]/main/div/div[2]/div[2]/div[1]/div[2]/div[2]/div[1]/p")
        print(f"\n{source_name} Entry Count for \"{search_term}\": {entry_count.text}")

        page_count = driver.find_element(By.XPATH, "//*[@id='rows-per-page']")
        page_count.click()
        page_count_value = driver.find_element(By.XPATH, "/html/body/div[9]/div[3]/ul/li[6]")
        page_count_value.click()
        print(f"{source_name} Set Entries Per Page to 100")

        time.sleep(3)
        body = driver.find_element(By.TAG_NAME, "body")
        goto_last_page_button = driver.find_element(By.CSS_SELECTOR, "button[aria-label='Go to last page']")

        FLAG = True

        while FLAG:
            # scroll to top of page by pressing home key
            body.send_keys(Keys.HOME)
            time.sleep(2)
            download_button = driver.find_element(By.XPATH, "//*[@id='basic-button']")
            download_button.click()
            export_excel_option = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[9]/div[3]/ul/li"))
            )
            export_excel_option.click()
            print(f"{source_name} Clicked Download Button to download Excel file")
            
            time.sleep(3)  # wait for download to complete

            wait = WebDriverWait(driver, 20)
            # --- Locate all down-arrow buttons ---
            arrows = wait.until(EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "button[aria-label='expand row']")
            ))

            print(f"Found {len(arrows)} rows")

            FIRST_TIME = True

            # --- Loop through each row ---
            for i, arrow in enumerate(arrows):
                try:
                    # Scroll into view and click arrow
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", arrow)
                    time.sleep(3)
                    # Use JavaScript click as a fallback if regular click fails
                    try:
                        arrow.click()
                    except Exception as e:
                        driver.execute_script("arguments[0].click();", arrow)
                    print(f"[{i+1}] Expanded row")

                    # Wait for PDF viewer to load (adjust selector if needed)
                    viewer_wait = WebDriverWait(driver, 120)
                    pdf_viewer = viewer_wait.until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, ".rpv-core__viewer")
                    ))
                    time.sleep(1)

                    if FIRST_TIME:
                        time.sleep(20)
                        FIRST_TIME = False
                    # Locate and click the download button with extended timeout
                    download_wait = WebDriverWait(driver, 120)  # Increase timeout to 120 seconds
                    download_btn = download_wait.until(EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "button[data-testid='get-file__download-button'], button[aria-label='Download']"))
                    )
                    download_btn.click()
                    print(f"[{i+1}] Download clicked")

                    time.sleep(5)  # Wait for download

                    arrow.click()
                    print(f"[{i+1}] Collapsed row")

                    time.sleep(2)
                except Exception as e:
                    print(f"Error in row {i+1}: {e}")
                    continue
            next_page_button = driver.find_element(By.CSS_SELECTOR, "button[aria-label='Go to next page']")
            # if next page button has disabled attribute, we are on last page
            if not next_page_button.get_attribute("disabled"):
                next_page_button.click()
                print(f"{source_name} Navigated to Next Page")
            goto_last_page_button = driver.find_element(By.CSS_SELECTOR, "button[aria-label='Go to last page']")
            if goto_last_page_button.get_attribute("disabled"):
                FLAG = False
    
        time.sleep(30)  # wait for download to complete

        process_results(search_term, source_name)
        
    except Exception as e:
        print(f"Error with {source_name}: {e}")
    finally:
        driver.quit()

def process_results(search_term, source_name):
    """Combine all Excel files in results folder into a single file."""
    # Sanitize search term for use as directory name
    search_term_dir = search_term.strip()
    # If starts with invalid char, prepend with underscore
    invalid_start_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|', '.']
    if search_term_dir and search_term_dir[0] in invalid_start_chars:
        search_term_dir = '_' + search_term_dir
    
    # Create directory structure: current_dir/search_term/source_name/
    base_dir = os.path.join(os.getcwd(), "results", search_term_dir, source_name)
    os.makedirs(base_dir, exist_ok=True)
    
    results_dir = base_dir
    
    # Get all .xls files in the results directory
    xls_files = [f for f in os.listdir(results_dir) if f.endswith('.xls')]
    
    if not xls_files:
        print("No .xls files found in results folder.")
        return
    
    print(f"Found {len(xls_files)} .xls files to process...")
    
    all_dataframes = []
    
    for xls_file in xls_files:
        file_path = os.path.join(results_dir, xls_file)
        try:
            # Since these are HTML files with .xls extension, use read_html
            tables = pd.read_html(file_path)
            if tables:
                # Get the first (and usually only) table from each file
                df = tables[0]
                all_dataframes.append(df)
                print(f"Processed: {xls_file} ({len(df)} rows)")
        except Exception as e:
            print(f"Error processing {xls_file}: {e}")
    
    if all_dataframes:
        # Combine all dataframes
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        
        # Remove duplicate rows if any
        original_count = len(combined_df)
        combined_df = combined_df.drop_duplicates()
        duplicates_removed = original_count - len(combined_df)
        
        # Save as a proper Excel file
        output_file = os.path.join(results_dir, f"{search_term}_{source_name}.xlsx")
        combined_df.to_excel(output_file, index=False, engine='openpyxl')
        
        print(f"\n{'='*50}")
        print(f"Combined {len(all_dataframes)} files into one")
        print(f"Total rows: {len(combined_df)}")
        if duplicates_removed > 0:
            print(f"Removed {duplicates_removed} duplicate rows")
        print(f"Output saved to: {output_file}")
        print(f"{'='*50}")
        
        # Delete the original .xls files after successful combination
        for xls_file in xls_files:
            file_path = os.path.join(results_dir, xls_file)
            try:
                os.remove(file_path)
                print(f"Deleted: {xls_file}")
            except Exception as e:
                print(f"Error deleting {xls_file}: {e}")
        
        print(f"Cleaned up {len(xls_files)} .xls files")
    else:
        print("No data to combine.")
    

def initialize_driver(search_term, source_name):
    """Initialize the Selenium WebDriver."""
    search_term_dir = search_term.strip()
    # If starts with invalid char, prepend with underscore
    invalid_start_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|', '.']
    if search_term_dir and search_term_dir[0] in invalid_start_chars:
        search_term_dir = '_' + search_term_dir
    
    # Create directory structure: current_dir/search_term/source_name/
    base_dir = os.path.join(os.getcwd(), "results", search_term_dir, source_name)
    os.makedirs(base_dir, exist_ok=True)
    
    download_directory = base_dir
    options = webdriver.ChromeOptions()
    # Set preferences for the download directory
    prefs = {
        "download.default_directory": download_directory,
        "download.prompt_for_download": False,  # Disable download prompt
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True  # Enable safe browsing
    }
    options.add_experimental_option("prefs", prefs)
    # options.add_argument('--headless')  # Run in headless mode
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver

def main():
    search_term = input("Enter search term: ").strip()
    
    print("\nSelect source:")
    print("1. LS")
    print("2. RS")
    print("3. Both")
    source_choice = input("Enter choice (1/2/3): ").strip()

    if source_choice in ['1', '3']:
        driver = initialize_driver(search_term, "LS")
        search_and_scrape(driver, LS_URL, search_term, "LS")
    
    if source_choice in ['2', '3']:
        driver = initialize_driver(search_term, "RS")
        search_and_scrape(driver, RS_URL, search_term, "RS")
    
    print("\n" + "="*50)
    print("Scraping completed!")

if __name__ == "__main__":
    main()