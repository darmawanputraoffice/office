from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import time
import os

def download_idx_financial_report(company_code, year, download_path='./downloads'):
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
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    
    try:
        print(f"Processing company code: {company_code}")

        # 1. Navigate to the company's profile page
        try:
            url = f"https://www.idx.co.id/id/perusahaan-tercatat/profil-perusahaan-tercatat/{company_code}"
            driver.get(url)
            time.sleep(1)
        except:
            print("Failed to load company profile page")
        
        # 2. Clicking 'Laporan Keuangan' tab
        try:
            laporan_keuangan = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//button[contains(., 'Laporan Keuangan')]"
                ))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", laporan_keuangan)
            driver.execute_script("arguments[0].click();", laporan_keuangan)
            time.sleep(1)
        except:
            print("Warning: Failed to click 'Laporan Keuangan' tab")

        # 3. Selecting year from dropdown
        try:
            comboboxes = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@role='combobox']"))
            )
            year_dropdown = None
            for idx, combobox in enumerate(comboboxes):
                try:
                    selected_text = combobox.find_element(By.CLASS_NAME, "vs__selected").text
                    if selected_text.isdigit() and len(selected_text) == 4:
                        year_dropdown = combobox
                        break
                except:
                    continue
                
            if year_dropdown:
                # Try typing in the input field
                try:
                    input_field = year_dropdown.find_element(By.TAG_NAME, "input")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_field)
                    input_field.click()
                    input_field.clear()
                    time.sleep(0.5)
                    input_field.send_keys(year)
                    time.sleep(1)
                    
                    from selenium.webdriver.common.keys import Keys
                    input_field.send_keys(Keys.ENTER)
                    time.sleep(1)
                        
                except Exception as e:
                    print(f"Warning: Failed to select the year input")
            else:
                print("Warning: Could not find year dropdown")
            
        except Exception as e:
            print(f"Warning: Failed to select the year input")
            import traceback
            traceback.print_exc()
        
        # 4. Looking for XLSX download link
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        xlsx_selectors = [
            f"//a[contains(@href, 'FinancialStatement-{year}-Tahunan-{company_code}.xlsx')]",
            f"//a[contains(@href, '{company_code}.xlsx') and contains(@href, '{year}')]",
            f"//a[contains(@href, '.xlsx') and contains(@href, '{year}')]",
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
                    break
            except:
                continue
        
        if xlsx_link:
            driver.execute_script("arguments[0].scrollIntoView(true);", xlsx_link)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", xlsx_link)
            time.sleep(1)
        else:
            print("Warning: XLSX download link not found")
        
    except Exception as e:
        print(f"ERROR processing {company_code}: {str(e)}")
        import traceback
        traceback.print_exc()
        
    finally:
        time.sleep(5)
        driver.quit()

def main():
    for company_code in ['MEDC']:
        for year in range(2025,2015,-1):
            download_idx_financial_report(company_code=company_code, year=year)

if __name__ == "__main__":
    main()