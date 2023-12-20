import pandas as pd
import os

def list_files_recursive(directory):
    all_files = []
    for root, dirs, files in os.walk(directory):
        all_files.extend([os.path.join(root, file) for file in files])
    return all_files

current_directory = os.getcwd()
files = list_files_recursive(current_directory)

tx_id_col_names = [
    "integratorTransId",
    "IntegratorTransId",
    "External Transaction Id",
    "External Payment Request â†’ Institution Trans ID",
    "Merchant Transaction Reference",
    "REMARKS2",
    "External Payment Request → Institution Trans ID",
    "Order Code",
    "Order ID",
    "Integrator Trans ID",
]


def getIdCol(df: pd.DataFrame):
    trans_id_col = ""
    for name in tx_id_col_names:
        if name in df.columns:
            trans_id_col = name
            break
    return trans_id_col
    
def getFailedTx():
    for file in files:
        if "Failed" in file:
            failed_tx_file = file
            break
    for file in files:
        if "Failed" not in file and "xlsx" in file:
            channel = os.path.basename(file)[:-15]
            confirmed_df = pd.read_excel(file)
            tx_id = getIdCol(confirmed_df)
            failed_df = pd.read_excel(failed_tx_file)
            confirmed_tx = (confirmed_df[tx_id].astype("string")).tolist()
            found = failed_df["Integrator Trans ID"].isin(confirmed_tx)
            with pd.ExcelWriter(failed_tx_file, engine="openpyxl", mode="a",if_sheet_exists="overlay") as writer:
                sheet_name = channel
                failed_df[found].to_excel(writer, sheet_name=sheet_name, index=False)


getFailedTx()
