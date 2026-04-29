import os
import time
import shutil
import glob
import pandas as pd
from selenium import webdriver
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException

root_dir=r"C:\Users\IlaBarshilia\ACME Barricades\FDOT Web Scraping - FDOT Web Scraping Data\Source Data"
base_path = root_dir
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
print(latest_folder_path)
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
    combined_df = pd.concat(all_dataframes, ignore_index=True)
    combined_df.columns = [col.replace('\n', '').strip() for col in combined_df.columns]
    combined_df=combined_df[['Contract','EstimateNumber', 'EstimateEnd Date','Project', 'Unit', 'Item Code', 'Previous', 
                             'This Est.', 'To-Date','Item Description']].copy().drop_duplicates()
    combined_df['Item Code'] = combined_df['Item Code'].str.strip().str.replace(r'\s+', ' ', regex=True)
    combined_df['Date_Downloaded'] = latest_row['date_str']
    
    print("All FDOT Downloads of Excel files have been combined into one file.")
else:
    print("No valid Excel files found to combine.")


### Pay items filter
Master_Pay_Item_List=pd.read_excel(root_dir + r"\\FDOT Master Pay Items.xlsx")
Master_Pay_Item_List.columns = Master_Pay_Item_List.iloc[1]
Master_Pay_Item_List = Master_Pay_Item_List.iloc[2:]
Master_Pay_Item_List = Master_Pay_Item_List[Master_Pay_Item_List['Acme Vlookup']== 'Include']
Master_Pay_Item_List['Item Number'] = Master_Pay_Item_List['Item Number'].str.strip().str.replace(r'\s+', ' ', regex=True)
# Create lowercased description columns for merging
# combined_df['Item Description Lower'] = combined_df['Item Description'].str.lower()
# Master_Pay_Item_List['Item Description Lower'] = Master_Pay_Item_List['Item Description'].str.lower()

filtered_df = combined_df.merge(
	Master_Pay_Item_List[['Item Number']],
	left_on=['Item Code'],
	right_on=['Item Number'],
	how='inner'
).drop(columns=['Item Number'])


### Filtered Data after latest cut off date
# Read Cut-off Dates
from datetime import datetime
cutoff_data = pd.read_excel(root_dir + r"\\FDOT Monthly CutOff Dates.xlsx")
current_month = (datetime.now().replace(day=1)).strftime('%B')
last_month = (datetime.now().replace(day=1) - pd.DateOffset(months=1)).strftime('%B')
last_cutoff_date = cutoff_data[cutoff_data['Month'] == last_month]['Cutoff date'].values[0]
current_cutoff_date = cutoff_data[cutoff_data['Month'] == current_month]['Cutoff date'].values[0]
# final_df = filtered_df[(filtered_df['EstimateEnd Date']>last_cutoff_date) & (filtered_df['EstimateEnd Date']<=current_cutoff_date)]
final_df = filtered_df[(filtered_df['EstimateEnd Date']>last_cutoff_date)]
yesterday = (datetime.today() - pd.DateOffset(days=1)).strftime('%Y-%b-%d')


### Append Final Acceptance Date to Final df
contracts_df_payment = pd.read_excel(root_dir + r"\\FDOT_All_Contracts_Latest.xlsx")
contracts_df_payment = contracts_df_payment[['Contract ID', 'Final Acceptance Date']].drop_duplicates()      
contracts_df_payment.rename(columns={"Contract ID": "Contract"}, inplace=True)        
contracts_df_payment['Final Acceptance Date'] = pd.to_datetime(contracts_df_payment['Final Acceptance Date']).dt.strftime('%m-%d-%Y')
final_df = final_df.merge(contracts_df_payment, how='left', on = "Contract")           
final_df = final_df.drop_duplicates()

final_df.to_excel(os.path.join(root_dir, 'FDOT_Output_Data_' + str(yesterday) + '.xlsx'), index=False)

final_df.to_excel(r"C:\Users\IlaBarshilia\ACME Barricades\FDOT Web Scraping - FDOT Web Scraping Data\Output Files\\" + str(current_month) +  '_' + str(datetime.now().strftime("%Y")) + '_Cutoff\\' + 'FDOT_Output_Data_' + str(yesterday) + '.xlsx', index=False)