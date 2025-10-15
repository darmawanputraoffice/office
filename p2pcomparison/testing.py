from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import time
import os

def download_idx_financial_report(company_code='BBCA', download_path='./downloads'):

    # Create download directory if it doesn't exist
    if not os.path.exists(download_path):
        os.makedirs(download_path)
    
    # Configure Chrome options
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": os.path.abspath(download_path),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Initialize the driver
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    
    try:
        print(f"Processing company code: {company_code}")
        
        # Step 1: Navigate directly to company profile page
        url = f"https://www.idx.co.id/id/perusahaan-tercatat/profil-perusahaan-tercatat/{company_code}"
        print(f"1. Navigating directly to {url}")
        driver.get(url)
        time.sleep(2)
        
        # Step 2: Click "Laporan Keuangan" tab/section
        print("2. Clicking 'Laporan Keuangan' tab...")
        laporan_keuangan = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//button[contains(., 'Laporan Keuangan')]"
            ))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", laporan_keuangan)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", laporan_keuangan)
        time.sleep(2)

        print("3. Selecting year 2024 from dropdown...")
        try:
            # Wait a bit for the page to fully load after clicking the tab
            time.sleep(2)
            
            # Find all v-select comboboxes (Vuetify dropdowns)
            comboboxes = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@role='combobox']"))
            )
            
            print(f"   Found {len(comboboxes)} combobox(es)")
            
            # Find the year dropdown - look for the one showing a 4-digit year
            year_dropdown = None
            for idx, combobox in enumerate(comboboxes):
                try:
                    selected_text = combobox.find_element(By.CLASS_NAME, "vs__selected").text
                    print(f"   Combobox {idx+1}: '{selected_text}'")
                    if selected_text.isdigit() and len(selected_text) == 4:  # It's a year
                        year_dropdown = combobox
                        print(f"   → Using this as year dropdown (currently: {selected_text})")
                        break
                except:
                    continue
            
            # Fallback: use second combobox
            if not year_dropdown and len(comboboxes) >= 2:
                year_dropdown = comboboxes[1]
                print(f"   Using second combobox as fallback")
            
            if year_dropdown:
                # Method 1: Try typing in the input field
                try:
                    # Find the input field inside the combobox
                    input_field = year_dropdown.find_element(By.TAG_NAME, "input")
                    
                    # Click on the input to focus it
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_field)
                    time.sleep(0.5)
                    input_field.click()
                    time.sleep(0.5)
                    
                    # Clear any existing text
                    input_field.clear()
                    time.sleep(0.3)
                    
                    # Type 2024
                    print("   Typing '2024' into dropdown...")
                    input_field.send_keys("2024")
                    time.sleep(1)
                    
                    # Press Enter or click on the matching option
                    from selenium.webdriver.common.keys import Keys
                    input_field.send_keys(Keys.ENTER)
                    print("   ✓ Pressed Enter")
                    time.sleep(2)
                    
                    # Verify selection
                    try:
                        selected = year_dropdown.find_element(By.CLASS_NAME, "vs__selected").text
                        print(f"   Current selection: {selected}")
                        if selected == "2024":
                            print("   ✓ Successfully selected 2024")
                        else:
                            print(f"   ⚠ Selection shows '{selected}', trying click method...")
                            raise Exception("Typing didn't work, try clicking")
                    except:
                        raise Exception("Typing didn't work, try clicking")
                        
                except Exception as e:
                    # Method 2: Fallback to clicking method
                    print(f"   Typing method failed, using click method...")
                    
                    # Click to open dropdown
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", year_dropdown)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", year_dropdown)
                    print("   Clicked dropdown to open")
                    time.sleep(2)
                    
                    # Find any visible listbox
                    listbox = None
                    try:
                        all_uls = driver.find_elements(By.TAG_NAME, "ul")
                        for ul in all_uls:
                            if ul.is_displayed() and ul.get_attribute('role') == 'listbox':
                                listbox = ul
                                print(f"   Found visible listbox")
                                break
                    except:
                        pass
                    
                    if listbox:
                        options = listbox.find_elements(By.TAG_NAME, "li")
                        print(f"   Found {len(options)} options in listbox")
                        
                        # Find and click 2024
                        for opt in options:
                            opt_text = opt.text.strip()
                            if opt_text == '2024':
                                print(f"   Clicking on 2024...")
                                driver.execute_script("arguments[0].click();", opt)
                                print("   ✓ Clicked on 2024")
                                time.sleep(2)
                                
                                # Verify selection
                                try:
                                    selected = year_dropdown.find_element(By.CLASS_NAME, "vs__selected").text
                                    print(f"   Current selection: {selected}")
                                except:
                                    pass
                                break
                    else:
                        print("   ⚠ Could not find listbox")
                        driver.save_screenshot(f"debug_listbox_{company_code}.png")
            else:
                print("   ⚠ Warning: Could not find year dropdown")
                driver.save_screenshot(f"debug_year_{company_code}.png")
            
        except Exception as e:
            print(f"   ✗ Error selecting year: {e}")
            driver.save_screenshot(f"debug_year_error_{company_code}.png")
            import traceback
            traceback.print_exc()
            print("   Continuing with default year...")
        
        # Step 4: Download the XLSX file
        print("4. Looking for XLSX download link...")
        
        # Scroll down to see all download options
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        
        # Look for the specific XLSX file pattern
        xlsx_selectors = [
            f"//a[contains(@href, 'FinancialStatement-2024-Tahunan-{company_code}.xlsx')]",
            f"//a[contains(@href, '{company_code}.xlsx') and contains(@href, '2024')]",
            "//a[contains(@href, '.xlsx') and contains(@href, '2024')]",
            "//a[contains(text(), 'XLSX') or contains(text(), 'xlsx')]",
            "//a[contains(@download, '.xlsx')]",
            "//button[contains(text(), 'XLSX')]",
            "//*[contains(@href, '.xlsx')]",
        ]
        
        xlsx_link = None
        for selector in xlsx_selectors:
            try:
                xlsx_link = driver.find_element(By.XPATH, selector)
                if xlsx_link:
                    print(f"   ✓ Found XLSX link using selector")
                    break
            except:
                continue
        
        if xlsx_link:
            # Scroll to the link and click
            driver.execute_script("arguments[0].scrollIntoView(true);", xlsx_link)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", xlsx_link)
            print(f"✓ Download initiated for {company_code}")
            print(f"   File: FinancialStatement-2024-Tahunan-{company_code}.xlsx")
            time.sleep(5)
        else:
            print(f"   ⚠ Warning: XLSX download link not found")
            print("\n   Available links on page:")
            all_links = driver.find_elements(By.TAG_NAME, "a")
            xlsx_count = 0
            for link in all_links:
                href = link.get_attribute("href")
                text = link.text
                if href and ('.xlsx' in href or '.xls' in href or 'download' in href.lower()):
                    print(f"      - {text}: {href}")
                    xlsx_count += 1
            
            if xlsx_count == 0:
                print("      - No XLSX links found")
            
            driver.save_screenshot(f"debug_download_{company_code}.png")
        
        print(f"\n{'='*60}")
        print(f"✓ Process completed for {company_code}")
        print(f"Files saved to: {os.path.abspath(download_path)}")
        
    except Exception as e:
        print(f"\n{'='*60}")
        print(f"ERROR processing {company_code}: {str(e)}")
        driver.save_screenshot(f"error_{company_code}.png")
        print(f"Screenshot saved as: error_{company_code}.png")
        import traceback
        traceback.print_exc()
        
    finally:
        print("\nClosing browser in 5 seconds...")
        time.sleep(5)
        driver.quit()

def download_multiple_reports(company_codes, download_path='./downloads'):
    """
    Download financial reports for multiple companies
    
    Args:
        company_codes: List of stock codes
        download_path: Path where files will be downloaded
    """
    print(f"\n{'='*60}")
    print(f"Starting batch download for {len(company_codes)} companies")
    print(f"{'='*60}\n")
    
    for i, code in enumerate(company_codes, 1):
        print(f"\n[{i}/{len(company_codes)}] Processing {code}...")
        download_idx_financial_report(code, download_path)
        
        if i < len(company_codes):
            print(f"\nWaiting 3 seconds before next download...")
            time.sleep(3)
    
    print(f"\n{'='*60}")
    print(f"✓ Batch download completed!")
    print(f"Total companies processed: {len(company_codes)}")
    print(f"{'='*60}")

def main():
    download_idx_financial_report()

if __name__ == "__main__":
    main()