import os
from datetime import datetime

# OVAs Rename Dictionary
OVAs_RENAMES = {
    "KB MTN COLLECTIONS": "KR MTN Debit",
    "KB MTN DISBURSEMENTS": "KR MTN Credit",
    "KB TELECEL COLLECTIONS": "KR Telecel Cashin",
    "KB TELECEL DISBURSEMENTS": "KR Telecel Cashout",
    "NPONTU COLLECTIONS": "Npontu MTN Credit",
    "NPONTU DISBURSEMENTS": "Npontu MTN Debit",
    "NGENIUS": "Ngenius",
    "GIP": "GIP",
    "MPGS": "MPGS",
}

# Function to convert the date format
def convert_date_format(date_str):
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    return date_obj.strftime("_%d %b_%y")

# Function to extract the date once and use it globally
def extract_date_from_filename(file_name):
    parts = file_name.split('_')
    if len(parts) >= 3:
        date_part = parts[1]  # Get the first date part
        return date_part
    return None

# Function to rename files and move them to the parent directory (cwd)
def rename_files(base_dir, renames_dict, global_date=None):
    parent_dir = os.path.dirname(base_dir)  # Get the parent directory (where the script is located)
    
    for dir_name, new_prefix in renames_dict.items():
        full_dir_path = os.path.join(base_dir, dir_name)
        
        if os.path.exists(full_dir_path):
            for file_name in os.listdir(full_dir_path):
                # Use the global date across all files if provided
                if global_date is not None:
                    date_to_use = global_date
                else:
                    # Extract the date from the file name for OVAs (KB MTN, NPONTU) files
                    if "_" in file_name and file_name.endswith(".csv"):
                        date_to_use = extract_date_from_filename(file_name)
                        if not date_to_use:
                            continue  # Skip if no valid date is found

                try:
                    # Handling files (ignoring time portion like 'T09' in the name)
                    formatted_date = convert_date_format(date_to_use)
                    new_name = f"{new_prefix}{formatted_date}.csv"

                    # Get full paths for the old file and the new location in the parent directory
                    old_file_path = os.path.join(full_dir_path, file_name)
                    new_file_path = os.path.join(parent_dir, new_name)  # Save to parent directory (cwd)

                    # Move and rename the file
                    os.rename(old_file_path, new_file_path)
                    print(f"Renamed and moved {file_name} to {new_name} in {parent_dir}")

                except Exception as e:
                    print(f"Error renaming file {file_name}: {e}")
        else:
            print(f"Directory {dir_name} does not exist")

# Run the function for OVAs first to extract the date
base_directory = os.getcwd()

# Rename files in the OVAs folder
ovas_dir = os.path.join(base_directory, 'OVAs')
global_mtn_date = None

for dir_name in OVAs_RENAMES:
    full_ovas_dir_path = os.path.join(ovas_dir, dir_name)
    if os.path.exists(full_ovas_dir_path):
        for file_name in os.listdir(full_ovas_dir_path):
            if "_" in file_name and file_name.endswith(".csv"):
                global_mtn_date = extract_date_from_filename(file_name)
                break  # Exit once a valid date is found
    if global_mtn_date:
        break  # Exit once a date is found

if global_mtn_date:
    # Rename OVAs files using the extracted date and move to the parent directory
    rename_files(ovas_dir, OVAs_RENAMES, global_mtn_date)
else:
    print("No valid date found in OVAs files.")

# Rename files in the mBase folder using the same global date and move to the parent directory
mbase_dir = os.path.join(base_directory, 'mBase')
rename_files(mbase_dir, mBase_RENAMES, global_mtn_date)
