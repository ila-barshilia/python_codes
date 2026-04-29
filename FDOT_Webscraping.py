import os
import time
import json
import shutil
import pandas as pd
from pathlib import Path
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    NoSuchElementException,
)
import sys
import re

# ================================
# CONFIG / CONSTANTS
# ================================
root_dir = r"C:\Users\acmefdotautomation\ACME Barricades\FDOT Web Scraping - FDOT Web Scraping Data\Source Data"
os.makedirs(root_dir, exist_ok=True)

STATE_DIR = os.path.join(root_dir, "state")
os.makedirs(STATE_DIR, exist_ok=True)
CHECKPOINT_PATH = os.path.join(STATE_DIR, "processed_contracts.json")

BASE_URL = "https://scoc.fdot.gov/#/active/1"
DOWNLOAD_PREFIX = "fdot_downloads_"
TODAY_STAMP = datetime.now().strftime("%Y-%b-%d")
TODAY_FOLDER_NAME = f"{DOWNLOAD_PREFIX}{TODAY_STAMP}"

# ================================
# HELPERS: STATE / CHECKPOINT
# ================================
def load_checkpoint(path: str) -> set:
    if not os.path.exists(path):
        return set()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return set(data if isinstance(data, list) else [])
    except Exception:
        # If corrupted, start fresh but keep the broken file as backup
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(path, f"{path}.bak_{ts}")
        except Exception:
            pass
        return set()

def save_checkpoint(path: str, processed: set):
    tmp = f"{path}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(sorted(list(processed)), f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

# ================================
# HELPERS: DOWNLOAD DIR SELECTION
# ================================
def ensure_today_download_dir(base_dir: str, today_folder_name: str) -> str:
    """
    Reuse today's folder if it exists; otherwise create it.
    This guarantees crash restarts do not create new folders.
    """
    download_dir = os.path.join(base_dir, today_folder_name)
    os.makedirs(download_dir, exist_ok=True)
    return download_dir

# ================================
# HELPERS: DISCOVER ALREADY DOWNLOADED CONTRACTS
# ================================
def discovered_contracts_in_dir(download_dir: str) -> set:
    """
    From the download folder:
      1) Parse 'Contract ID - {id}.xlsx' filenames
      2) For any .xlsx without that pattern, peek the first row Contract value
    """
    found = set()
    excel_files = [f for f in os.listdir(download_dir) if f.lower().endswith((".xlsx", ".xlsm"))]

    pat = re.compile(r'^Contract ID -\s*(?P<cid>[A-Za-z0-9\-]+)\.xlsx$', re.IGNORECASE)

    for fn in excel_files:
        m = pat.match(fn)
        fp = os.path.join(download_dir, fn)

        if m:
            found.add(m.group("cid").strip())
            continue

        # Otherwise, try to peek its Contract column
        try:
            df = pd.read_excel(fp, engine="openpyxl")
            if "Contract" in df.columns:
                val = str(df.loc[0, "Contract"]).strip()
                if val and val.lower() != "nan":
                    found.add(val)
        except Exception:
            # ignore unreadable partials; they will be retried
            pass

    return found

# ================================
# HELPERS: DOWNLOAD COMPLETION
# ================================
def wait_for_download_and_get_path(download_dir: str, before_files: set, timeout: int = 240) -> str:
    """
    Waits until a new file appears in download_dir and it is fully downloaded
    (no .crdownload and size stable). Returns the path to the completed file.
    """
    end_time = time.time() + timeout
    last_size = {}

    while time.time() < end_time:
        current = set(os.listdir(download_dir))
        new_files = list(current - before_files)
        # Skip non-excel or temp downloads
        new_files = [f for f in new_files if not f.lower().endswith(".crdownload")]

        # If we see a new file, confirm it has stable size
        for fn in new_files:
            fp = os.path.join(download_dir, fn)
            if not os.path.isfile(fp):
                continue

            size = os.path.getsize(fp)
            if fn in last_size and last_size[fn] == size and size > 0:
                # Stable between loops
                return fp
            last_size[fn] = size

        time.sleep(1.2)

    raise TimeoutError("Timed out waiting for download to complete.")

def rename_to_contract_id(file_path: str, contract_id: str) -> str:
    """
    Rename the downloaded file to 'Contract ID - {contract_id}.xlsx'
    """
    folder = os.path.dirname(file_path)
    new_name = f"Contract ID - {contract_id}.xlsx"
    new_path = os.path.join(folder, new_name)

    # If a file already exists with same name, replace it (idempotent resume)
    if os.path.abspath(file_path) != os.path.abspath(new_path):
        if os.path.exists(new_path):
            try:
                os.remove(new_path)
            except Exception:
                # fallback: unique suffix
                base, ext = os.path.splitext(new_name)
                new_path = os.path.join(folder, f"{base}__dup_{int(time.time())}{ext}")
        os.replace(file_path, new_path)
    return new_path

# ================================
# STALE/FRAME-SAFE HELPERS FOR UI
# ================================
def reset_to_root(driver):
    """Safely reset to the top-level browsing context."""
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

def get_contract_rows(driver):
    """Return a fresh list of row elements under the Contracts table."""
    return driver.find_elements(By.CSS_SELECTOR, 'table[aria-label="Contracts"] tbody tr')

def iterate_contract_ids_on_page(driver):
    """
    Returns the list of contract IDs visible on the current page (fresh).
    Assumes the 3rd <td> has the contract ID (as in your table).
    """
    ids = []
    rows = get_contract_rows(driver)
    for r in rows:
        tds = r.find_elements(By.TAG_NAME, 'td')
        if len(tds) >= 3:
            cid = (tds[2].text or "").strip()
            if cid:
                ids.append(cid)
    return ids

def open_contract_detail_by_id(driver, wait, contract_id):
    """
    Locate the details link for a contract (by ID) and open it.
    Uses fresh locators to avoid stale references.
    """
    reset_to_root(driver)
    row_xpath = f'//table[@aria-label="Contracts"]//tbody//tr[.//td[3][normalize-space()="{contract_id}"]]'
    link_xpath = f'{row_xpath}//a[@aria-label="View Contract Details"]'

    row = wait.until(EC.presence_of_element_located((By.XPATH, row_xpath)))
    link = wait.until(EC.element_to_be_clickable((By.XPATH, link_xpath)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
    driver.execute_script("arguments[0].click();", link)

    # Wait for details view to be present
    wait.until(EC.presence_of_element_located(
        (By.XPATH, '//a[contains(text(), "Back to Active Contract List")]')
    ))

def run_excel_report(driver, wait, contract_id, download_dir):
    """
    Opens the report modal (if available), selects Excel, runs the report,
    waits for the download to complete, and returns the filepath.
    Returns None if no report or no Excel or timeout.
    """
    reset_to_root(driver)

    # Open report modal button for this contract
    report_button_xpath = f'//*[contains(@title, "Get estimate detail report for {contract_id}")]'
    report_buttons = driver.find_elements(By.XPATH, report_button_xpath)
    if not report_buttons:
        print(f"[{contract_id}] No estimate detail report available. Skipping.")
        return None

    report_button = report_buttons[0]
    driver.execute_script("arguments[0].click();", report_button)

    # Wait modal
    wait.until(EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "modal-content")]')))

    # Format dropdown → Excel
    format_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, '//select')))
    options = format_dropdown.find_elements(By.TAG_NAME, 'option')
    excel_option = next((o for o in options if "Excel" in (o.text or "")), None)
    if not excel_option:
        print(f"[{contract_id}] Excel option not found. Skipping.")
        return None

    driver.execute_script("""
        arguments[0].selected = true;
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, excel_option)

    # Optional: wait overlay to go away if present
    try:
        wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'div.overlay')))
    except TimeoutException:
        pass

    # Run report
    before_files = set(os.listdir(download_dir))
    run_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//button[contains(@title, "Run Report - Get Estimate Detail")]')
    ))
    driver.execute_script("arguments[0].click();", run_button)

    # Wait for the file to land and stabilize
    try:
        fp = wait_for_download_and_get_path(download_dir, before_files, timeout=240)
        return fp
    except TimeoutError:
        print(f"[{contract_id}] Download timed out.")
        return None

def close_modal_if_open(driver):
    """Best-effort to close the modal if it is open."""
    try:
        reset_to_root(driver)
        cancel_button = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@title="Cancel" and contains(@class, "close")]'))
        )
        driver.execute_script("arguments[0].click();", cancel_button)
        time.sleep(0.3)
    except Exception:
        pass

def back_to_list(driver, wait):
    """Navigate back to the list view reliably."""
    try:
        reset_to_root(driver)
        back_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//a[contains(text(), "Back to Active Contract List")]')
        ))
        driver.execute_script("arguments[0].click();", back_button)
        wait.until(EC.presence_of_element_located((By.XPATH, '//table[@aria-label="Contracts"]')))
        time.sleep(0.5)
    except Exception:
        driver.get(BASE_URL)
        wait.until(EC.presence_of_element_located((By.XPATH, '//table[@aria-label="Contracts"]')))

def click_next_page(driver, wait) -> bool:
    """
    Attempts to click the 'next page' button with retries on stale.
    Returns False if we're at the last page or cannot proceed.
    """
    reset_to_root(driver)
    for _ in range(3):
        try:
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@title="next page"]')))
            cls = (btn.get_attribute("class") or "")
            if "disabled" in cls:
                print("Next button is disabled. Reached last page.")
                return False
            driver.execute_script("arguments[0].click();", btn)
            wait.until(EC.presence_of_element_located((By.XPATH, '//table[@aria-label="Contracts"]')))
            time.sleep(0.6)
            return True
        except StaleElementReferenceException:
            time.sleep(0.3)
            continue
        except TimeoutException:
            print("Next Page not found/clickable (timeout).")
            return False
    print("Next Page not found/clickable after retries (stale).")
    return False

# ================================
# LOAD INPUT DATA (same as yours, small tweaks)
# ================================
contracts_df_latest = pd.read_excel(os.path.join(root_dir, "FDOT_All_Contracts_Latest.xlsx"))
contracts_df_latest = contracts_df_latest[['Contract ID', 'Status']].copy().drop_duplicates()
contracts_df_latest = contracts_df_latest[contracts_df_latest['Status'] == 'FINAL PAYMENT MADE']
contracts_df_latest = contracts_df_latest.rename(columns={'Contract ID': 'Contract'})

contracts_df_last_month = pd.read_excel(os.path.join(root_dir, "FDOT_All_Contracts_Combined_Till_Last_Month.xlsx"))
contracts_df_last_month = contracts_df_last_month[['Contract ID', 'Status']].copy().drop_duplicates()
contracts_df_last_month = contracts_df_last_month[contracts_df_last_month['Status'] == 'FINAL PAYMENT MADE'].drop(columns=['Status'])
contracts_df_last_month = contracts_df_last_month.rename(columns={'Contract ID': 'Contract'})

common_contracts = pd.merge(contracts_df_latest, contracts_df_last_month, on='Contract', how='inner')
common_contracts = common_contracts[['Contract','Status']].copy().drop_duplicates()

# Master Job List
Master_Job_List = pd.read_excel(os.path.join(root_dir, "Acme DOT Jobs (from Sage 100).xlsx"))
Master_Job_List.columns = Master_Job_List.iloc[0]  # Set first row as header
Master_Job_List = Master_Job_List.iloc[1:]
Master_Job_List = Master_Job_List[Master_Job_List['JobType'] == 'DOT'].copy()
master_jobnos = set(Master_Job_List["JobNo"].astype(str).str.strip().tolist())

# ================================
# PREP: Download directory + checkpoint + already downloaded
# ================================
download_dir = ensure_today_download_dir(root_dir, TODAY_FOLDER_NAME)

processed_contracts = load_checkpoint(CHECKPOINT_PATH)
already_downloaded = discovered_contracts_in_dir(download_dir)
already_done = processed_contracts.union(already_downloaded)

# Build target set: in Master list, NOT in final payment made, NOT already done
final_paid = set(common_contracts["Contract"].astype(str).str.strip().tolist())
targets = {c for c in master_jobnos if c not in final_paid and c not in already_done}

print(f"Contracts to process (remaining): {len(targets)}")

# ================================
# SELENIUM RUNNER (stale-safe + id-based navigation)
# ================================
def run_scrape_once(targets: set):
    """
    One browser session that walks pages and downloads for 'targets'.
    Raises exceptions to caller; outer runner handles retry/backoff.
    """
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    options.add_experimental_option("prefs", prefs)
    # Optional: headless
    # options.add_argument("--headless=new")
    options.page_load_strategy = "eager"

    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(0.3)  # small implicit wait; rely on explicit waits
    wait = WebDriverWait(driver, 25)

    try:
        driver.get(BASE_URL)
        wait.until(EC.presence_of_element_located((By.XPATH, '//table[@aria-label="Contracts"]')))

        page_count = 1
        while True:
            print(f"\n--- Processing Page {page_count} ---\n")
            reset_to_root(driver)
            wait.until(EC.presence_of_element_located((By.XPATH, '//table[@aria-label="Contracts"]')))

            # Freshly read visible contract IDs on this page
            visible_ids = iterate_contract_ids_on_page(driver)

            for contract_id in visible_ids:
                if contract_id not in targets:
                    continue

                print(f"\nProcessing Contract ID: {contract_id}")
                try:
                    # Open details view
                    open_contract_detail_by_id(driver, wait, contract_id)

                    # Try report flow
                    try:
                        fp = run_excel_report(driver, wait, contract_id, download_dir)
                        if fp:
                            try:
                                new_path = rename_to_contract_id(fp, contract_id)
                                print(f"[{contract_id}] Downloaded and renamed to: {os.path.basename(new_path)}")
                            except Exception as re_err:
                                print(f"[{contract_id}] Rename failed: {re_err}")
                            processed_contracts.add(contract_id)
                            save_checkpoint(CHECKPOINT_PATH, processed_contracts)
                            targets.discard(contract_id)
                        else:
                            # Either no report or timeout/Excel missing
                            processed_contracts.add(contract_id)
                            save_checkpoint(CHECKPOINT_PATH, processed_contracts)
                            targets.discard(contract_id)

                    except Exception as e:
                        print(f"[{contract_id}] Failed within report modal: {e}")

                    # Cleanup UI and go back to list
                    close_modal_if_open(driver)
                    back_to_list(driver, wait)

                except StaleElementReferenceException:
                    print(f"[{contract_id}] Stale while opening details; reloading list.")
                    driver.get(BASE_URL)
                    wait.until(EC.presence_of_element_located((By.XPATH, '//table[@aria-label="Contracts"]')))
                    time.sleep(0.5)
                    continue
                except Exception as e:
                    print(f"[{contract_id}] Unexpected error: {e}. Recovering to list.")
                    try:
                        driver.save_screenshot(os.path.join(download_dir, f"error_{contract_id}.png"))
                    except Exception:
                        pass
                    driver.get(BASE_URL)
                    wait.until(EC.presence_of_element_located((By.XPATH, '//table[@aria-label="Contracts"]')))
                    time.sleep(0.5)
                    continue

            # Pagination (stale-safe)
            proceeded = click_next_page(driver, wait)
            if not proceeded:
                break
            page_count += 1

        print("Pagination complete.")
    finally:
        try:
            driver.quit()
        except Exception:
            pass

def self_healing_runner(targets: set, max_restarts: int = 50):
    """
    Keeps restarting the browser session if it crashes, until targets are exhausted
    or max_restarts reached. Uses exponential backoff.
    """
    if not targets:
        print("No remaining contracts to process. Skipping scrape.")
        return

    restart = 0
    while targets and restart <= max_restarts:
        try:
            run_scrape_once(targets)
        except Exception as e:
            wait_s = min(60, 2 ** min(restart, 5))  # exponential backoff up to 60s
            print(f"[SESSION ERROR] {e}. Restarting in {wait_s}s (attempt {restart+1}/{max_restarts})...")
            time.sleep(wait_s)
            restart += 1
        else:
            if targets:
                print(f"{len(targets)} targets remain; restarting a new browser session to continue...")
                restart += 1
                time.sleep(2)
                continue
            break

# Run the self-healing scraper
self_healing_runner(targets)

# ================================
# POST-PROCESSING (uses today's folder deterministically)
# ================================
print(f"Using download folder: {download_dir}")
all_dataframes = []

for filename in os.listdir(download_dir):
    if filename.lower().endswith(('.xlsx', '.xlsm')):
        file_path = os.path.join(download_dir, filename)
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            df.columns = [col.replace('\n', '').strip() for col in df.columns]
            # Attempt to identify contract
            contract_id = str(df.loc[0, "Contract"]).strip() if "Contract" in df.columns else ""
            if not contract_id or contract_id.lower() == "nan":
                print(f"Skipping {filename}: Contract ID missing")
                continue

            # Ensure renamed format (idempotent)
            desired_name = f"Contract ID - {contract_id}.xlsx"
            desired_path = os.path.join(download_dir, desired_name)
            if os.path.basename(file_path) != desired_name:
                try:
                    if os.path.exists(desired_path):
                        os.remove(desired_path)
                    os.rename(file_path, desired_path)
                    file_path = desired_path
                except Exception as e:
                    print(f"Rename during combine skipped ({filename}): {e}")

            if not df.empty:
                all_dataframes.append(df)
            else:
                print(f"Skipped empty file: {filename}")
        except Exception as e:
            print(f"Error reading {filename}: {e}")

if all_dataframes:
    combined_df = pd.concat(all_dataframes, ignore_index=True)
    combined_df.columns = [c.replace('\n', '').strip() for c in combined_df.columns]
    # Ensure columns exist before selection
    expected_cols = ['Contract','EstimateNumber', 'EstimateEnd Date','Project', 'Unit',
                     'Item Code', 'Previous', 'This Est.', 'To-Date','Item Description']
    available_cols = [c for c in expected_cols if c in combined_df.columns]
    combined_df = combined_df[available_cols].copy().drop_duplicates()
    if 'Item Code' in combined_df.columns:
        combined_df['Item Code'] = combined_df['Item Code'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
    combined_df['Date_Downloaded'] = TODAY_STAMP
    combined_df = combined_df.drop_duplicates()
    print("All FDOT Downloads of Excel files have been combined into one file.")
else:
    print("No valid Excel files found to combine.")
    combined_df = pd.DataFrame(columns=['Contract','EstimateNumber', 'EstimateEnd Date','Project', 'Unit',
                                        'Item Code', 'Previous', 'This Est.', 'To-Date','Item Description',
                                        'Date_Downloaded'])

# ================================
# Pay items filter (same as yours)
# ================================
Master_Pay_Item_List = pd.read_excel(os.path.join(root_dir, "FDOT Master Pay Items.xlsx"))
Master_Pay_Item_List.columns = Master_Pay_Item_List.iloc[1]
Master_Pay_Item_List = Master_Pay_Item_List.iloc[2:]
Master_Pay_Item_List = Master_Pay_Item_List[Master_Pay_Item_List['Acme Vlookup']== 'Include']
Master_Pay_Item_List['Item Number'] = Master_Pay_Item_List['Item Number'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)

filtered_df = combined_df.merge(
    Master_Pay_Item_List[['Item Number']],
    left_on=['Item Code'],
    right_on=['Item Number'],
    how='inner'
).drop(columns=['Item Number'])

# -------------------- Today & strings --------------------
today = datetime.today().date()
yesterday = (datetime.today() - pd.DateOffset(days=1)).strftime('%Y-%b-%d')

# -------------------- Load cutoffs --------------------
cutoff_path = Path(root_dir) / "FDOT Monthly CutOff Dates.xlsx"
cutoff_df = pd.read_excel(cutoff_path)

# Normalize dates and build a (year, month_num) -> cutoff_date map
cutoff_df['Cutoff date'] = pd.to_datetime(cutoff_df['Cutoff date'], errors='coerce').dt.date
if cutoff_df['Cutoff date'].isna().any():
    sys.exit("Cutoff sheet has unparseable 'Cutoff date' values.")

cutoff_df['year']      = pd.to_datetime(cutoff_df['Cutoff date']).dt.year
cutoff_df['month_num'] = pd.to_datetime(cutoff_df['Cutoff date']).dt.month
cutoff_map = {(int(r.year), int(r.month_num)): r['Cutoff date'] for _, r in cutoff_df.iterrows()}

# -------------------- Helpers --------------------
def month_anchors(d: date):
    first_curr = d.replace(day=1)
    first_prev = first_curr - relativedelta(months=1)
    return (first_prev.year, first_prev.month, first_curr.year, first_curr.month)

def in_inclusive_window(day: date, start: date) -> bool:
    return start <= day <= (start + timedelta(days=14))

def select_active_cutoff(d: date, lookup: dict):
    py, pm, cy, cm = month_anchors(d)
    prev_cutoff = lookup.get((py, pm))
    curr_cutoff = lookup.get((cy, cm))

    if prev_cutoff and in_inclusive_window(d, prev_cutoff):
        return {'active_cutoff': prev_cutoff, 'period_year': py, 'period_month': pm}
    if curr_cutoff and in_inclusive_window(d, curr_cutoff):
        return {'active_cutoff': curr_cutoff, 'period_year': cy, 'period_month': cm}
    return None

# -------------------- Decide whether to run today --------------------
decision = select_active_cutoff(today, cutoff_map)
if decision is None:
    sys.exit(f"No valid run window for today {today}. Exiting without output.")

active_cutoff_date = decision['active_cutoff']
period_year        = decision['period_year']
period_month       = decision['period_month']

period_start = date(period_year, period_month, 1)

# -------------------- Validate input dataframe --------------------
if 'filtered_df' not in globals():
    sys.exit("Variable 'filtered_df' is not defined. Prepare it before running this script.")

required_cols = {'EstimateEnd Date', 'Contract'}
missing = required_cols - set(filtered_df.columns)
if missing:
    sys.exit(f"filtered_df missing required columns: {missing}")

filtered_df = filtered_df.copy()
filtered_df['EstimateEnd Date'] = pd.to_datetime(filtered_df['EstimateEnd Date'], errors='coerce')

# -------------------- Apply filter (LOWER BOUND ONLY) --------------------
mask = (filtered_df['EstimateEnd Date'] >= pd.to_datetime(period_start))
final_df = filtered_df.loc[mask].copy()

if final_df.empty:
    sys.exit(
        f"No rows found using lower-bound-only filter "
        f"(EstimateEnd Date >= {period_start}) for active cutoff {active_cutoff_date}. Exiting."
    )

# -------------------- Append Final Acceptance Date --------------------
contracts_path = Path(root_dir) / "FDOT_All_Contracts_Latest.xlsx"
contracts_df_payment = pd.read_excel(contracts_path)
contracts_df_payment = (
    contracts_df_payment[['Contract ID', 'Final Acceptance Date']]
    .drop_duplicates()
    .rename(columns={"Contract ID": "Contract"})
)
contracts_df_payment['Final Acceptance Date'] = pd.to_datetime(
    contracts_df_payment['Final Acceptance Date'], errors='coerce'
).dt.strftime('%m-%d-%Y')

final_df = final_df.merge(contracts_df_payment, how='left', on="Contract").drop_duplicates()

# -------------------- Outputs --------------------
file_name = f"FDOT_Output_Data_{yesterday}.xlsx"
out_path_1 = Path(root_dir) / file_name

target_month_name = period_start.strftime('%B')
target_year_num   = period_start.year

out_dir_2 = (
    Path(r"C:\Users\acmefdotautomation\ACME Barricades\FDOT Web Scraping - FDOT Web Scraping Data\Output Files")
    / f"{target_month_name}_{target_year_num}_Cutoff"
)
out_path_2 = out_dir_2 / file_name

Path(root_dir).mkdir(parents=True, exist_ok=True)
out_dir_2.mkdir(parents=True, exist_ok=True)

final_df.to_excel(out_path_1, index=False)
final_df.to_excel(out_path_2, index=False)

print(
    f"Run OK | Active cutoff: {active_cutoff_date} | Lower bound: >= {period_start} | "
    f"Rows: {len(final_df)} | Output: {out_path_1} and {out_path_2}"
)