import pandas as pd
import os

#list all the files in the directory
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

#get the id column used for a particular file
def getIdCol(df: pd.DataFrame):
    trans_id_col = ""
    for name in tx_id_col_names:
        if name in df.columns:
            trans_id_col = name
            break
    return trans_id_col
    
def getFailedTx():
    #locating the Failed transactions file
    for file in files:
        if "Failed" in file:
            #store data from file in this variable
            failed_tx_file = file
            break
    for file in files:
        #Looking through all the other files
        if "Failed" not in file and "xlsx" in file:
            #getting the channel from the name of the file
            channel = os.path.basename(file)[:-15]

            #putting the OVA file into its own dataframe
            confirmed_df = pd.read_excel(file)

            #getting the id col for the OVA file
            tx_id = getIdCol(confirmed_df)

            #reading the failed transaction data
            failed_df = pd.read_excel(failed_tx_file)

            #Put all ids from OVA in a list
            confirmed_tx = (confirmed_df[tx_id].astype("string")).tolist()

            #Compare the ids in the failed tx dataframe with the ones in the list. Results are stored in "found" as a list
            found = failed_df["Integrator Trans ID"].isin(confirmed_tx)

            #write the data in "found" into the Failed Transactions File in the specified sheet
            with pd.ExcelWriter(failed_tx_file, engine="openpyxl", mode="a",if_sheet_exists="overlay") as writer:
                sheet_name = channel
                failed_df[found].to_excel(writer, sheet_name=sheet_name, index=False)

#calling the function to begin the program
getFailedTx()
