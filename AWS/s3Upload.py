import boto3
import json
import pandas as pd
import io
import os

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
month_abbr_as_list = month_to_number.keys()
months_as_list = month_abbreviations.values()
s3_client = boto3.client("s3")
buckets = s3_client.list_buckets()

MTN_UNWANTED_HEADERS = [
    "Provider Category",
    "Information",
    "Initiated By",
    "On Behalf Of",
    "External Amount",
    "Currency",
    "External FX Rate",
    "External Service Provider",
    "Discount",
    "Promotion",
    "Coupon",
    "Balance",
]

VODAFONE_UNWANTED_HEADERS = ["Reason Type", "Opposite Party", "Linked Transaction ID"]


def getData(file: str):
    if "Vodafone" in file:
        unwanted_headers = VODAFONE_UNWANTED_HEADERS
        skiprows = 5
        channel = "Vodafone"
    elif "MTN" in file:
        unwanted_headers = MTN_UNWANTED_HEADERS
        skiprows = 0
        channel = "MTN"
    elif "QUIPU" in file:
        unwanted_headers = []
        skiprows = 0
        channel = "CardProvider"
    elif "MPGS" in file:
        unwanted_headers = []
        skiprows = 0
        channel = "CardProvider"

    try:
        data = pd.read_csv(file, skiprows=skiprows)
    except:
        try:
            data = pd.read_excel(file, skiprows=skiprows)
        except Exception as e:
            print(f"###########################{e}#########################")
            pass
    data = data.drop(columns=unwanted_headers)
    print(data)
    return data, channel


def list_files_recursive(directory):
    all_files = []
    for root, dirs, files in os.walk(directory):
        all_files.extend([os.path.join(root, file) for file in files])
    return all_files

current_directory = os.getcwd()
files = list_files_recursive(f"{current_directory}/AWS/data")


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

files = removeDupFile()
for file in files:
    if "Ops" in file or "DS_Store" in file or "Statement" in file or "Utility" in file:
        continue
    if "BANK" in file or "KOWRI" in file:
        print(f"--------------------------------{file}-------------------------")
        file_df, channel = getData(file=file)
        os.remove(file)
        file_df.to_csv(file, index=False)
        file_name = os.path.basename(file)
        year = f"20{file_name[-6:-4]}"
        month = file_name[-10:-7]
        path_month = month_to_number[file_name[-10:-7]]
        day = file_name[-13:-10]
        if "_" in day:
            day = day.replace("_", "")
            path_day = f"0{day}"
        else:
            path_day = day

        if "MTN Debit" in file_name:
            collOrDisb = "Collection"
            new_file_name = (
                f"KB_MOMO_MTN_Collection_{year}_{path_month}_{path_day}.csv"
            )
        elif "MTN Credit" in file_name:
            new_file_name = (
                f"KB_MOMO_MTN_Disbursement_{year}_{path_month}_{path_day}.csv"
            )
            collOrDisb = "Disbursement"
        elif "Vodafone Credit" in file_name:
            new_file_name = (
                f"KB_MOMO_VODAFONE_Collection_{year}_{path_month}_{path_day}.csv"
            )
            collOrDisb = "Collection"
        elif "Vodafone Debit" in file_name:
            new_file_name = (
                f"KB_MOMO_VODAFONE_Disbursement_{year}_{path_month}_{path_day}.csv"
            )
            collOrDisb = "Disbursement"
        elif "QUIPU" in file_name:
            new_file_name = (
                f"KB_CARD_CAL_Transactions_{year}_{path_month}_{path_day}.csv"
            )
            collOrDisb = "CAL"
        elif "MPGS" in file_name:
            new_file_name = (
                f"KB_CARD_GT_Transactions_{year}_{path_month}_{path_day}.csv"
            )
            collOrDisb = "GTBANK"

        Key = f"KowriBusiness/{channel}/{collOrDisb}/year={year}/month={path_month}/day={path_day}/{new_file_name}"
        s3_client.upload_file(file, "kowribusiness-datalake", Key)
