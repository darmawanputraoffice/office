from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
import time
import csv
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os

def generate_date_range(start_date="2025-09", end_date="2000-01"):
    dates = []
    current = datetime.strptime(start_date, "%Y-%m")
    end = datetime.strptime(end_date, "%Y-%m")
    
    while current >= end:
        dates.append(current.strftime("%Y-%m"))
        current = current - relativedelta(months=1)
    
    return dates

def wait_for_table_update(driver, wait, timeout=10):
    """Wait for table to finish loading/updating"""
    try:
        # Wait for any loading indicators to disappear
        time.sleep(0.5)
        
        # Wait for table body to be present and have content
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody")))
        
        # Additional check: wait for at least one row or no-data message
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
                if rows and len(rows) > 0:
                    # Check if row has actual content
                    first_row_text = rows[0].text.strip()
                    if first_row_text and first_row_text != "":
                        return True
                time.sleep(0.2)
            except StaleElementReferenceException:
                time.sleep(0.2)
                continue
        
        return True
    except TimeoutException:
        return False

def extract_bottom_layer_headers(driver):
    """Extract all table headers"""
    headers = ["Company_Code", "Period"]
    
    try:
        # Get all th elements from the table
        header_elements = driver.find_elements(By.CSS_SELECTOR, "thead th")
        
        for header in header_elements:
            header_text = header.text.strip()
            if header_text:
                headers.append(header_text)
        
        print(f"Extracted {len(headers)-2} column headers")
        
    except Exception as e:
        print(f"Warning: Could not extract headers: {e}")
    
    return headers

def scrape_idx_data_to_csv(company_code="MEDC", start_date="2025-09", end_date="2000-01", output_file="idx_data.csv"):
    # Initialize Chrome driver
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    # Uncomment to run headless
    # options.add_argument('--headless')
    
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)  # Increased timeout
    
    # Generate date range
    date_list = generate_date_range(start_date, end_date)
    print(f"Will scrape {len(date_list)} periods from {start_date} to {end_date}")
    
    # Track if this is first write (to write headers)
    is_first_write = not os.path.exists(output_file)
    csv_file = open(output_file, 'a', newline='', encoding='utf-8')
    csv_writer = None
    headers = None
    
    try:
        # Step 1: Navigate to the URL (only once)
        url = "https://www.idx.co.id/id/data-pasar/laporan-statistik/digital-statistic/monthly/financial-report-and-ratio-of-listed-companies/financial-data-and-ratio?filter=eyJ5ZWFyIjoiMjAyMyIsIm1vbnRoIjoiOSIsInF1YXJ0ZXIiOjAsInR5cGUiOiJtb250aGx5In0%3D"
        print(f"\nNavigating to: {url}")
        driver.get(url)
        
        # Wait for initial page load
        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Cari laporan...']")))
        time.sleep(2)  # Give page time to fully render
        
        # Process each period
        for idx, month_year in enumerate(date_list, 1):
            print(f"\n{'='*60}")
            print(f"Processing {idx}/{len(date_list)}: {company_code} - {month_year}")
            print(f"{'='*60}")
            
            try:
                # Step 2: Fill in the "Cari laporan..." field with company code
                print(f"Searching for company: {company_code}")
                search_field = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Cari laporan...']"))
                )
                search_field.click()
                time.sleep(0.1)
                search_field.clear()
                time.sleep(0.1)
                search_field.send_keys(company_code)
                time.sleep(0.3)
                
                # Step 3: Fill in the "Bulan & Tahun" field (clear first)
                print(f"Setting date: {month_year}")
                date_field = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Bulan & Tahun']"))
                )
                
                # Clear the field completely
                date_field.click()
                time.sleep(0.1)
                date_field.send_keys(Keys.CONTROL + "a")
                time.sleep(0.1)
                date_field.send_keys(Keys.BACKSPACE)
                time.sleep(0.1)
                
                # Type the new date
                date_field.send_keys(month_year)
                time.sleep(0.3)
                
                # Step 4: Click the "Terapkan" button
                print("Clicking 'Terapkan' button...")
                apply_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Terapkan')]"))
                )
                apply_button.click()
                
                # Wait for table to update with new data
                print("Waiting for data to load...")
                if not wait_for_table_update(driver, wait):
                    print("Warning: Table may not have loaded properly")
                
                # Additional wait to ensure data is fully rendered
                time.sleep(1)
                
                # Step 5: Extract the data
                print("Extracting data...")
                
                # Wait for table to be present
                table = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table"))
                )
                
                # Extract headers (only on first iteration or if not yet extracted)
                if headers is None:
                    headers = extract_bottom_layer_headers(driver)
                    
                    # Initialize CSV writer with headers
                    csv_writer = csv.writer(csv_file)
                    if is_first_write:
                        csv_writer.writerow(headers)
                        print(f"CSV Headers: {headers}")
                        is_first_write = False
                
                # Extract table rows
                row_elements = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
                
                if not row_elements or len(row_elements) == 0:
                    print(f"No data found for {month_year}")
                    # Write a row indicating no data
                    no_data_row = [company_code, month_year] + ["NO DATA"] * (len(headers) - 2)
                    csv_writer.writerow(no_data_row)
                    csv_file.flush()
                    continue
                
                rows_written = 0
                for row in row_elements:
                    try:
                        cells = row.find_elements(By.CSS_SELECTOR, "td")
                        row_data = [company_code, month_year]
                        
                        for cell in cells:
                            cell_text = cell.text.strip()
                            row_data.append(cell_text)
                        
                        # Only write if we have actual data
                        if len(row_data) > 2 and any(cell != "" for cell in row_data[2:]):
                            csv_writer.writerow(row_data)
                            rows_written += 1
                    except StaleElementReferenceException:
                        print(f"Warning: Stale element, skipping row")
                        continue
                
                csv_file.flush()  # Ensure data is written immediately
                print(f"✓ Successfully scraped and saved {rows_written} rows for {month_year}")
                
                # Small delay between requests to be respectful
                time.sleep(0.5)
                
            except Exception as e:
                print(f"✗ Error processing {month_year}: {str(e)}")
                # Write error row
                if csv_writer and headers:
                    error_row = [company_code, month_year] + [f"ERROR: {str(e)}"] + [""] * (len(headers) - 3)
                    csv_writer.writerow(error_row)
                    csv_file.flush()
                
                # Take screenshot for debugging
                screenshot_name = f"error_{company_code}_{month_year.replace('-', '_')}.png"
                try:
                    driver.save_screenshot(screenshot_name)
                    print(f"Screenshot saved as {screenshot_name}")
                except:
                    pass
                
                # Continue to next period
                continue
        
        print(f"\n{'='*60}")
        print(f"✓ Scraping completed! Data saved to: {output_file}")
        print(f"{'='*60}")
        
    except Exception as e:
        print(f"\n✗ Fatal error: {str(e)}")
        try:
            driver.save_screenshot("fatal_error.png")
        except:
            pass
        
    finally:
        csv_file.close()
        print("\nClosing browser...")
        driver.quit()


if __name__ == "__main__":
    # Configuration
    company_codes = ["SOCI"] # ["PGAS", "AKRA", "ENRG", "ELSA"]
    start_date = "2025-09"
    end_date = "2020-01"
    for code in company_codes:
        output_file = f"idx_data_{code}.csv"
        scrape_idx_data_to_csv(code, start_date, end_date, output_file)