import boto3
import json
import pandas as pd
import io
import os

# Dictionary to map month abbreviations to numbers
month_to_number = {
    "Jan": "01",
    "Feb": "02",
    "Mar": "03",
    "Apr": "04",
    "May": "05",
    "Jun": "06",
    "Jul": "07",
    "Aug": "08",
    "Sep": "09",
    "Oct": "10",
    "Nov": "11",
    "Dec": "12",
}

# Dictionary to map month abbreviations to full names
month_abbreviations = {
    "Jan": "January",
    "Feb": "February",
    "Mar": "March",
    "Apr": "April",
    "May": "May",
    "Jun": "June",
    "Jul": "July",
    "Aug": "August",
    "Sep": "September",
    "Oct": "October",
    "Nov": "November",
    "Dec": "December",
}

# List of month abbreviations
month_abbr_as_list = month_to_number.keys()
# List of month names
months_as_list = month_abbreviations.values()

# Create an S3 client
s3_client = boto3.client("s3")
# List all S3 buckets
buckets = s3_client.list_buckets()

# List of unwanted headers for MTN transactions
MTN_UNWANTED_HEADERS = [
    "Provider category",
    # "Information",
    "Initiated by",
    "On behalf of",
    "External amount",
    # "Currency",
    "Currency.1",
    "Currency.2",
    "Currency.3",
    "Currency.4",
    "Currency.5",
    "Currency.6",
    "Currency.7",
    "Currency.8",
    "Currency.9",
    "Currency.10",
    "Currency.11",
    "Currency.12",
    "Currency.13",
    "Currency.15",
    "Currency.16",
    "External FX rate",
    "External service provider",
    "Discount",
    "From / Promotion",
    "To / Promotion",
    "Coupon",
    "Balance",
]

# List of unwanted headers for Vodafone transactions
VODAFONE_UNWANTED_HEADERS = ["Reason Type", "Opposite Party", "Linked Transaction ID"]


# Function to get data from a file
def getData(file: str):
    # Determine unwanted headers and skiprows based on the file type
    if "Telecel" in file:
        unwanted_headers = VODAFONE_UNWANTED_HEADERS
        skiprows = 5
        channel = "Vodafone"
    elif "MTN" in file:
        unwanted_headers = MTN_UNWANTED_HEADERS
        skiprows = 0
        channel = "MTN"
    elif "QUIPU" in file or "MPGS" in file:
        unwanted_headers = []
        skiprows = 0
        channel = "CardProvider"
    elif "Utility" in file:
        unwanted_headers = []
        skiprows = 5
        channel = "Utility"
    elif "Ngenius" in file:
        unwanted_headers = []
        skiprows = 0
        channel = "CardProvider"

    try:
        print("--------------------------- attempting csv read ---------------------------------")
        # Try reading the file as CSV
        data = pd.read_csv(file, skiprows=skiprows)
    except:
        try:
            print("--------------------------- attempting excel read ---------------------------------")
            # If reading as CSV fails, try reading as Excel
            data = pd.read_excel(file, skiprows=skiprows)
            print("--------------------------- read data -------------------------------------")
        except Exception as e:
            # If both methods fail, print an error and continue to the next file
            print(f"###########################{e}#########################")
            pass

    # Drop unwanted headers from the DataFrame
    data = data.drop(columns=unwanted_headers)
    print(data)
    return data, channel


# Function to list all files in a directory and its subdirectories
def list_files_recursive(directory):
    all_files = []
    for root, dirs, files in os.walk(directory):
        all_files.extend([os.path.join(root, file) for file in files])
    return all_files


# Get the current working directory
current_directory = os.getcwd()
# List all files in the "data" directory and its subdirectories
files = list_files_recursive(f"{current_directory}\\data")
print(files)


# Function to remove duplicate files
def removeDupFile():
    try:
        for file in files:
            indices_to_drop = [
                i
                for i, item in enumerate(files)
                if os.path.basename(file)[:-4] in item and "xlsx" in item
            ]

            if indices_to_drop:
                files.pop(indices_to_drop[0])
    except:
        for file in files:
            indices_to_drop = [
                i
                for i, item in enumerate(files)
                if os.path.basename(file)[:-4] in item and "csv" in item
            ]

            if indices_to_drop:
                files.pop(indices_to_drop[0])
    return files


# List of files with errors during processing
error_list = []
# Remove duplicate files
# files = removeDupFile()
for file in files:
    try:
        # Skip files with specific keywords
        if "DS_Store" in file or "Statement" in file:
            continue
        else:
            print(f"--------------------------------{file}-------------------------")
            # Get data from the file and determine the channel
            file_df, channel = getData(file=file)
            # Remove the original file
            os.remove(file)
            # Save the modified DataFrame to a CSV file
            file_df.to_csv(file, index=False)
            file_name = os.path.basename(file)
            if 'csv' in file_name:
                year = f"20{file_name[-6:-4]}" #csv 
                month = file_name[-10:-7] #csv
                path_month = month_to_number[file_name[-10:-7]] #csv
                day = file_name[-13:-10] #csv
            elif 'xlsx' in file_name:
                year = f"20{file_name[-7:-5]}" #xlsx
                month = file_name[-11:-8] #xlsx
                path_month = month_to_number[file_name[-11:-8]] #xlsx
                day = file_name[-14:-11] #xlsx
            if "_" in day:
                day = day.replace("_", "")
                path_day = f"0{day}"
            else:
                path_day = day
            channel_type = "KowriBusiness"

            # Determine the type of transaction and create a new file name
            if "KR MTN Debit" in file_name:
                collOrDisb = "MTN-GH-Collections"
                new_file_name = (
                    f"KB_MOMO_MTN_Collection_{year}_{path_month}_{path_day}.csv"
                )
            elif "KR MTN Credit" in file_name:
                collOrDisb = "MTN-GH-Disbursements"
                new_file_name = (
                    f"KB_MOMO_MTN_Disbursement_{year}_{path_month}_{path_day}.csv"
                )
            elif "Telecel Cashin" in file_name:
                collOrDisb = "Vodafone-GH-Collections"
                new_file_name = (
                    f"KB_MOMO_VODAFONE_Collection_{year}_{path_month}_{path_day}.csv"
                )
            elif "Telecel Cashout" in file_name:
                collOrDisb = "Vodafone-GH-Disbursements"
                new_file_name = (
                    f"KB_MOMO_VODAFONE_Disbursement_{year}_{path_month}_{path_day}.csv"
                )
            elif "QUIPU" in file_name:
                collOrDisb = "Card-GH-CALQUIPU"
                new_file_name = (
                    f"KB_CARD_CAL_Transactions_{year}_{path_month}_{path_day}.csv"
                )
            elif "Ngenius" in file_name:
                collOrDisb = "Card-GH-NGENIUS"
                new_file_name = (
                    f"NGENIUS_{year}_{path_month}_{path_day}.csv"
                )
            elif "MPGS" in file_name:
                collOrDisb = "Card-GH-GTMPGS"
                new_file_name = (
                    f"KB_CARD_GT_Transactions_{year}_{path_month}_{path_day}.csv"
                )
            elif "Npontu MTN Debit" in file_name:
                collOrDisb = "SecurePay-GH-Disbursements"
                channel_type = "KowriPartner"
                new_file_name = (
                    f"SecurePay_Disbursements_{year}_{path_month}_{path_day}.csv"
                )
            elif "Npontu MTN Credit" in file_name:
                channel_type = "KowriPartner"
                collOrDisb = "SecurePay-GH-Collections"
                new_file_name = (
                    f"SecurePay_Collections_{year}_{path_month}_{path_day}.csv"
                )
            elif "Utility" in file_name:
                collOrDisb = "Both"
                new_file_name = (
                    f"KB_MOMO_VODAFONE_Utility_{year}_{path_month}_{path_day}.csv"
                )

            # Specify the S3 key for the new file
            Key = f"{channel_type}/{collOrDisb}/year={year}/month={path_month}/day={path_day}/{new_file_name}"
            # Upload the file to the S3 bucket
            s3_client.upload_file(file, "all-kowri-datalake", Key)
    except KeyError as e:
        # If a KeyError occurs, add the file to the error list and continue to the next file
        error_list.append(file)
        continue

# Print the list of files with errors
print("-----------------------------------------------------------------------------------------------------------------------------------------")
print(
    f"These files need to be rechecked and uploaded. Please confirm the naming or check if the headers to be removed are available: \n{error_list}"
)
print("\n\n !!!! SCRIPT COMPLETED !!!!")