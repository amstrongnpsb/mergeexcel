import os
import pandas as pd
from datetime import datetime


# Set the folder path where the Excel files are located
folder_path = './excel'

# Create an empty list to store the data from each Excel file
all_data = []

# Loop through all files in the folder
for file_name in os.listdir(folder_path):
    # Check if the file is an Excel file
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        # Construct the full file path
        file_path = os.path.join(folder_path, file_name)
        
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)
        
        # Append the DataFrame to the list
        all_data.append(df)

# Concatenate all DataFrames in the list into a single DataFrame
merged_data = pd.concat(all_data, ignore_index=True)
now = datetime.now()
output_folder_path = './result'
current_time = now.strftime("%d_%m_%Y_%H_%M_%S")
# Write the merged DataFrame to a new Excel file
output_file = os.path.join(output_folder_path, f"po_tracking_export_all_{current_time}.xlsx")
merged_data.to_excel(output_file, index=False)
print(f"All Excel files have been merged into '{output_file}'")