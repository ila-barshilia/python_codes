import pandas as pd
import os

# Set the folder path containing CSV files
folder_path = r'C:\Users\IlaBarshilia\Downloads\2025 SBN Reports'

# List to hold individual DataFrames
df_list = []

# Loop through all CSV files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.csv'):
        file_path = os.path.join(folder_path, filename)
        
        # Read the CSV file
        df = pd.read_csv(file_path)
        
        # Extract text after the last '-' in the filename (excluding extension)
        label = filename.rsplit('-', 1)[-1].replace('.csv', '')
        
        # Add a new column with the extracted label
        df['Month - Year'] = label + ' - 2025'
        
        # Append to the list
        df_list.append(df)

# Combine all DataFrames into one
final_df = pd.concat(df_list, ignore_index=True)

# Optional: Save the combined DataFrame
final_df.to_csv('Output_SBN_YTD.csv', index=False)