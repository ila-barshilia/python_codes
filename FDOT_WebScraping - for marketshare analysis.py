import os
import time
import pandas as pd
pd.set_option('future.no_silent_downcasting', True)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
from dateutil.relativedelta import relativedelta

# Setup download directory
# root_dir = input("Enter the folder path where 'fdot_downloads' should be created: ")
root_dir = r"C:\Users\IlaBarshilia\ACME Barricades\FDOT Web Scraping - FDOT Web Scraping Data\FDOT Market Share Analysis"
download_dir = os.path.join(root_dir, "fdot_downloads_marketshare_" + datetime.now().strftime("%Y-%b-%d"))
os.makedirs(download_dir, exist_ok=True)

# Setup Chrome WebDriver
options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_setting_values.automatic_downloads": 1
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

processed_contracts = set()

try:
    driver.get("https://scoc.fdot.gov/#/active/1")
    wait.until(EC.presence_of_element_located((By.XPATH, '//table[@aria-label="Contracts"]')))

    page_count = 1

    while True:
        # print(f"\n--- Processing Page {page_count} ---\n")

        wait.until(EC.presence_of_element_located((By.XPATH, '//table[@aria-label="Contracts"]')))
        rows = driver.find_elements(By.CSS_SELECTOR, 'table[aria-label="Contracts"] tbody tr')

        row_index = 0
        while row_index < len(rows):
            row = rows[row_index]
            row_index += 1

            cells = row.find_elements(By.TAG_NAME, "td")
            if not cells or all(cell.text.strip() == "" for cell in cells):
                continue  # Skip blank rows silently

            contract_id = cells[2].text.strip()
            is_active = cells[0].text.strip().lower() == "yes"

            # if contract_id in Downloaded_Job_List["Contract"].values:
            #     print(f"\nSkipping Contract ID: {contract_id} — found in Downloaded job list.")
            #     continue
            
            # ✅ Only process if in master list, not in common_contracts and not already processed
            if contract_id in processed_contracts:
                print(f"\nSkipping Contract ID: {contract_id} — already processed.")
                continue

            processed_contracts.add(contract_id)
            print(f"\nProcessing Contract ID: {contract_id}")

            try:
                link = cells[2].find_element(By.XPATH, './/a[@aria-label="View Contract Details"]')
                driver.execute_script("arguments[0].click();", link)
                time.sleep(1.5)

                wait.until(EC.presence_of_element_located((
                    By.XPATH,
                    '//div[contains(@class, "contract-detail")] | //a[contains(text(), "Back to Active Contract List")]')))

                try:
                    report_button_xpath = f'//*[contains(@title, "Get estimate detail report for {contract_id}")]'
                    report_buttons = driver.find_elements(By.XPATH, report_button_xpath)

                    if not report_buttons:
                        print(f"[{contract_id}] No estimate detail report available. Skipping.")
                    else:
                        report_button = report_buttons[0]
                        driver.execute_script("arguments[0].click();", report_button)
                        time.sleep(1)

                        wait.until(EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "modal-content")]')))
                        # driver.save_screenshot(f"modal_{contract_id}.png")

                        format_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, '//select')))
                        excel_option = None
                        for option in format_dropdown.find_elements(By.TAG_NAME, 'option'):
                            if "Excel" in option.text:
                                excel_option = option
                                break

                        if excel_option:
                            driver.execute_script("""
                                arguments[0].selected = true;
                                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                            """, excel_option)
                            print(f"[{contract_id}] Excel format selected.")
                            time.sleep(1)
                        else:
                            print(f"[{contract_id}] Excel option not found.")
                            continue

                        try:
                            wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'div.overlay')))
                        except TimeoutException:
                            print(f"[{contract_id}] Overlay did not disappear in time.")

                        before_files = set(os.listdir(download_dir))
                        run_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(@title, "Run Report - Get Estimate Detail")]')))
                        driver.execute_script("arguments[0].click();", run_button)
                        print(f"[{contract_id}] Triggered Excel download.")

                        time.sleep(30)

                        # Click on cancel button after triggering the download
                        # cancel_button = driver.find_element(By.XPATH, '//button[@title="Cancel" and contains(@class, "close")]')
                        # driver.execute_script("arguments[0].click();", cancel_button)


                        try:
                            cancel_button = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, '//button[@title="Cancel" and contains(@class, "close")]')))
                            driver.execute_script("arguments[0].click();", cancel_button)
                            print(f"[{contract_id}] Cancel button clicked.")
                            time.sleep(10)
                        except (TimeoutException, NoSuchElementException) as e:
                            pass

                except Exception as e:
                    print(f"[{contract_id}] Failed to download report: {e}")

                back_button = driver.find_element(By.XPATH, '//a[contains(text(), "Back to Active Contract List")]')
                driver.execute_script("arguments[0].click();", back_button)
                wait.until(EC.presence_of_element_located((By.XPATH, '//table[@aria-label="Contracts"]')))
                time.sleep(2)
                rows = driver.find_elements(By.CSS_SELECTOR, 'table[aria-label="Contracts"] tbody tr')

            except Exception as e:
                print(f"[{contract_id}] Error during processing: {e}")
                driver.save_screenshot(f"error_{contract_id}.png")
                driver.get("https://scoc.fdot.gov/#/active/1")
                time.sleep(2)
                break

        # Try to go to the next page
        try:
            next_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@title="next page"]')))
            if "disabled" in next_button.get_attribute("class"):
                print("Next button is disabled. Reached last page.")
                break
            driver.execute_script("arguments[0].click();", next_button)
            time.sleep(2)
            page_count += 1
        except TimeoutException:
            print("Next Page button not found or not clickable. Ending pagination. All contracts processed.")
            break

finally:
    driver.quit()

base_path=root_dir
folder_paths = [os.path.join(base_path, name) 
                for name in os.listdir(base_path) 
                if os.path.isdir(os.path.join(base_path, name)) and "fdot_downloads" in name]
# Create a DataFrame from folder_paths
df_folders = pd.DataFrame({'folder_path': folder_paths})
# Extract the folder name (after last '\')
df_folders['folder_name'] = df_folders['folder_path'].apply(lambda x: os.path.basename(x))
# Extract the date using regex
df_folders['date_str'] = df_folders['folder_name'].str.extract(r'(\d{4}-[A-Za-z]{3}-\d{2})')
# Convert to datetime for comparison
df_folders['date'] = pd.to_datetime(df_folders['date_str'], format='%Y-%b-%d', errors='coerce')
# Find the row with the latest date
latest_row = df_folders.loc[df_folders['date'].idxmax()]
latest_folder_path = latest_row['folder_path']
print('latest folder path: ', latest_folder_path)

all_dataframes = []
for filename in os.listdir(latest_folder_path):
    if filename.endswith('.XLSX') or filename.endswith('.xlsx'):
        file_path = os.path.join(latest_folder_path, filename)
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            if not df.empty:
                all_dataframes.append(df)
            else:
                print(f"Skipped empty file: {filename}")
        except Exception as e:
            print(f"Error reading {filename}: {e}")

if all_dataframes:
    combined_df1 = pd.concat(all_dataframes, ignore_index=True)
    combined_df1.columns = [col.replace('\n', '').strip() for col in combined_df1.columns]
    combined_df1.rename(columns={'This Est..1': 'This Est. Amount'}, inplace=True)
    combined_df1=combined_df1[['Contract', 'Vendor Name', 'EstimateNumber', 'EstimateEnd Date','Project', 'Unit', 'Item Code', 'Previous', 'This Est. Amount', 'AmountTo-Date'
                               , 'This Est.', 'To-Date','Item Description']].copy().drop_duplicates()
    combined_df1['Item Code'] = combined_df1['Item Code'].str.strip().str.replace(r'\s+', ' ', regex=True)
    
    print("All FDOT Downloads of Excel files have been combined into one file.")
else:
    print("No valid Excel files found to combine.")

today = datetime.today()
target_month = today - relativedelta(months=2)
end_of_month = pd.Timestamp(target_month).replace(day=1) + pd.offsets.MonthEnd(0)
combined_df1=combined_df1.drop_duplicates()
combined_df1 = combined_df1[combined_df1['EstimateEnd Date']<= end_of_month]
combined_df1.to_csv(root_dir + "\\Overall_FDOT.csv", index=False)

pay_items = pd.read_excel(r"C:\Users\IlaBarshilia\ACME Barricades\FDOT Web Scraping - FDOT Web Scraping Data\Source Data\FDOT Master Pay Items.xlsx")
pay_items.columns = pay_items.iloc[1]
pay_items = pay_items.iloc[2:]
pay_items = pay_items[pay_items['Acme Vlookup'] == 'Include']
pay_items['Item Number'] = pay_items['Item Number'].str.strip().str.replace(r'\s+', ' ', regex=True)
pay_items = pay_items[['Item Number']]
pay_items = pay_items.drop_duplicates()

combined_df2 = combined_df1.merge(pay_items, left_on='Item Code', right_on='Item Number', how='inner').drop(columns="Item Number")
combined_df2 = combined_df2.replace({r'[\r\n]+': ' '}, regex=True)
combined_df2.to_csv(root_dir + "\\Overall_FDOT_Relevant_Pay_Items_Only.csv", index=False)