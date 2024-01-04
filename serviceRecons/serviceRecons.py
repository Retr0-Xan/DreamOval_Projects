import pandas as pd
from typing import Literal, Tuple
from datetime import datetime
from datetime import date
from datetime import timedelta
import os, sys

tx_id_col_names = [
    "integratorTransId",
    "IntegratorTransId",
    "Transaction Id",
    "TransId",
    "External Transaction Id",
    "BillerTransId",
    "Id",
    "External Payment Request â†’ Institution Trans ID",
    "Merchant Transaction Reference",
    "REMARKS2",
    "External Payment Request → Institution Trans ID",
    "Order Code",
    "Order ID",
    "Integrator Trans ID",
    "CP-ID",
    "OP-ID",
    "Payment Reference",
]
amount_col_names = [
    "Amount",
    "amount",
    "Paid in",
    "Paid In",
    "Withdrawn",
    "AMOUNT",
    "Amount ($)",
    "Actual Amount ($)",
    "Transaction Amount (GHC.)",
    "TRANSACTION_AMOUNT",
    "Transaction Amount (amount only)",
    "Order Amount (amount only)",
    "Real Total",
    "Amount(Cedis)",
    "Amount(GHC)",
    "Amount(Old Cedis)",
]

status_col_names = [
    "Status",
    "Transaction Status",
]

success_conditions = ["SUCCESS", "Success", "Successful", "CONFIRMED"]

bundle_codes = {
    "DATABUNDBMXDLNR": 3,
    "DATABUNDGAP1_NR": 5,
    "BDLBTBOSSUD1_NR": 10,
    "DATABUND2MO_NR1": 10,
    "BDLBTBOSSUWE": 15,
    "DATABUND2MO_NR2": 20,
    "DATABUNDGAP4_NR": 20,
    "BDLBTBOSSUD2_NR": 43.5,
    "DATABUND2MO_NR3": 50,
    "DATABUND100NR": 100,
    "DATABUNDJSTRENR": 200,
    "DATABUNDJMAX_NR": 300,
    "DATABUNDJBOptNR": 400,
    "DATANVSTRDLY": 0.5,
    "DATANVCHTDLY": 2,
    "DATANVDBDL1": 10,
    "DATANVDBDL2": 20,
    "DATANVDBDL3": 50,
    "DATANVDBDL4": 100,
    "BDLDATABNIGHT2": 2,
    "BDLHRBOOST2": 1,
    "BDLHRBOOST3": 2,
    "BDLDATABNIGHT3": 3,
    "DATANVDR5WLY": 5,
    "DATANVDBDL5": 200,
    "DATANVDBDL6": 300,
    "DATANVDBDL7": 400,
    "DATANVDR1DLY": 1,
}


# Function to list all files in a directory and its subdirectories
def list_files_recursive(directory):
    all_files = []
    for root, dirs, files in os.walk(directory):
        all_files.extend([os.path.join(root, file) for file in files])
    return all_files


# Get the current working directory
current_directory = os.getcwd()
# List all files in the "data" directory and its subdirectories
files = list_files_recursive(f"{current_directory}/SERVICES")

mtn_airtime_files = []
mtn_data_files = []
voda_airtime_files = []
voda_data_files = []

for file in files:
    if ".DS_Store" in file:
        continue
    if "MTN Airtime" in file:
        mtn_airtime_files.append(file)
        if "EXPORT" in mtn_airtime_files[0]:
            break
        else:
            mtn_airtime_files.insert(0,mtn_airtime_files.pop(1))
    elif "MTN Data" in file:
        mtn_data_files.append(file)
    elif "Vodafone Airtime" in file:
        voda_airtime_files.append(file)
    elif "Vodafone Data" in file:
        voda_data_files.append(file)


def run_recons(
    dataList: list, num_lines_of_header: Tuple[int, int], file_output_name: str
):
    for file in dataList:
        if "KB" not in file:
            try:
                ova_df = pd.read_csv(file, skiprows=num_lines_of_header[0])
            except:
                ova_df = pd.read_excel(file, skiprows=num_lines_of_header[0])
        else:
            try:
                int_df = pd.read_csv(file, skiprows=num_lines_of_header[1])
            except:
                int_df = pd.read_excel(file, skiprows=num_lines_of_header[1])

    status_col = ""
    for name in status_col_names:
        if name in ova_df.columns:
            if not ova_df[name].isna().all():
                status_col = name
                break

    ova_df = ova_df.loc[ova_df[status_col].isin(success_conditions)]
    print(ova_df)
    print(int_df)
    ova_volume = len(ova_df)
    int_volume = len(int_df)

    print(f"{file_output_name} ova volume : {ova_volume}")
    print(f"{file_output_name} int volume : {int_volume}")

    amount_col = ""
    for name in amount_col_names:
        if name in ova_df.columns:
            if not ova_df[name].isna().all():
                amount_col = name
                break
    ova_value = 0.0
    if "KB Vodafone Data" in dataList[0] or "KB Vodafone Data" in dataList[1]:
        for code in ova_df[amount_col]:
            ova_value += bundle_codes[code]
    else:
        ova_value = (ova_df[amount_col].astype("float")).abs().sum()

    print(f"{file_output_name} OVA_VALUE : {ova_value}")

    for name in amount_col_names:
        if name in int_df.columns:
            if not int_df[name].isna().all():
                amount_col = name
                break

    int_value = (int_df[amount_col].astype("float")).abs().sum()
    print(f"{file_output_name} INT_VALUE : {int_value}")


run_recons(mtn_data_files, num_lines_of_header=[0, 0], file_output_name="MTN DATA")
run_recons(voda_data_files, num_lines_of_header=[0, 0], file_output_name="VODA Data")
run_recons(mtn_airtime_files, num_lines_of_header=[0, 0], file_output_name="MTN Airtime")
# run_recons(voda_data_files, num_lines_of_header=[0, 0], file_output_name="VODA Data")
