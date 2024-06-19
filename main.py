import os
import pandas as pd
from datetime import datetime

# Set the folder path where the Excel files are located
folder_path = './excel'

# Create an empty list to store the data from each Excel file
all_data = []

# Function to read an Excel file and append the DataFrame to the list
def read_excel_file(file_path):
    df = pd.read_excel(file_path)
    all_data.append(df)

# Function to display initial processing message
def initial_processing_message():
    print("Reading Excel files...")

# Function to display processing completed message
def merged_processing_message():
    print("Merging files...")
# Start the initial processing message
initial_processing_message()
# Loop through all files in the folder
for file_name in os.listdir(folder_path):
    # Check if the file is an Excel file
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        # Construct the full file path
        file_path = os.path.join(folder_path, file_name)
        
        # Read the Excel file into a DataFrame
        read_excel_file(file_path)

# Display the processing completed message
merged_processing_message()

# Concatenate all DataFrames in the list into a single DataFrame
merged_data = pd.concat(all_data, ignore_index=True)
now = datetime.now()
output_folder_path = './result'
current_time = now.strftime("%d_%m_%Y_%H_%M_%S")
# Write the merged DataFrame to a new Excel file
output_file = os.path.join(output_folder_path, f"po_tracking_export_all_{current_time}.xlsx")
merged_data.to_excel(output_file, index=False)
print(f"All Excel files have been merged into '{output_file}'")