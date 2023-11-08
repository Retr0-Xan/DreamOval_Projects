# version 2.0.0
import pandas as pd
from typing import Literal, Tuple

# from dateTime import yesterday, GIPdate, current_month, recons_yesterday
from datetime import datetime
from datetime import date
from datetime import timedelta
import os

tx_id_col_names = ["integratorTransId", "IntegratorTransId","Transaction Id","TransId","External Transaction Id","BillerTransId","Id","External Payment Request â†’ Institution Trans ID","Merchant Transaction Reference","REMARKS2","External Payment Request → Institution Trans ID",'Order Code','Order ID']
amount_col_names = ["Amount", "amount","Paid in","Paid In","Withdrawn","AMOUNT",'Amount ($)',"Actual Amount ($)","Transaction Amount (GHC.)","TRANSACTION_AMOUNT",'Transaction Amount (amount only)',"Order Amount (amount only)",'Real Total']
def check_for_file(file_name):
    if file_name in os.listdir():
        return file_name
    else:
        return None
    
def get_ova_id_col(df: pd.DataFrame):
    id_names = ["integratorTransId", "IntegratorTransId","Transaction Id","TransId","External Transaction Id","BillerTransId","Id","External Payment Request â†’ Institution Trans ID","Merchant Transaction Reference","REMARKS2","External Payment Request → Institution Trans ID"]
    for id in id_names:
        if id in df:
            return id
def get_int_id_col(df: pd.DataFrame):
    id_names = ["integratorTransId", "IntegratorTransId","Transaction Id","TransId","External Transaction Id","BillerTransId","Id","External Payment Request â†’ Institution Trans ID","Merchant Transaction Reference","REMARKS2","External Payment Request → Institution Trans ID"]
    for id in id_names:
        if id in df:
            return id

def get_date(
    date: date, format: Literal["gip"] | Literal["normal"] | Literal["recons"]
):
    """
    Formats a given date according to the specified format.

    Parameters:
    date (date): A datetime.date object representing the date to be formatted.
    format (Literal["gip"] | Literal["normal"] | Literal["recons"]): The format to use for formatting the date.
        - "gip": Format date as "YYYYMMDD" (e.g., 20231201).
        - "recons": Format date as "YYYY-MM-DD" (e.g., 2023-12-01).
        - "normal": Format date as day_month_year (e.g., 1_Jan_23).

    Returns:
    str: The formatted date string based on the specified format.

    Example:
    input_date = date(2023, 12, 1)

    # Format as "YYYYMMDD"
    gip_format = get_date(input_date, "gip")
    print(gip_format)
    # Output: "20231201"

    # Format as "YYYY-MM-DD"
    recons_format = get_date(input_date, "recons")
    print(recons_format)
    # Output: "2023-12-01"

    # Format as "normal"
    normal_format = get_date(input_date, "normal")
    print(normal_format)
    # Output: "1 Jan_23"

    Note:
    - This function allows formatting a date in three different formats based on the 'format' parameter.
    - 'gip' format represents the date as "YYYYMMDD".
    - 'recons' format represents the date as "YYYY-MM-DD".
    - 'normal' format represents the date as day_month_year (e.g., 1_Jan_23).
    """
    if format == "gip":
        return date.strftime("%Y%m%d")  # Output date in format: 20231201
    elif format == "recons":
        return date.strftime("%Y-%m-%d")
    else:
        return date.strftime("%-d %b_%y")  # Output date in format: 1_Jan_23

def get_write_double_ova_val(
        ova_files: Tuple[str | None, str | None],
        num_lines_of_header: Tuple[int, int],
        alt_recons_name: str,
        file_output_name: str
        ):
    try:
        ova_df1 = pd.read_excel(ova_files[0],skiprows=num_lines_of_header[0])
        ova_df2 = pd.read_excel(ova_files[1],skiprows=num_lines_of_header[1])

        ova_volume = len(ova_df1) + len(ova_df2)
        recons_file = (
            f"{alt_recons_name} - Recons.xlsx"
            if ova_files[0] is None
            else f"{ova_files[0][:-5]} - Recons.xlsx"
        )
        df1_amount_col = ""
        for name in amount_col_names:
            if name in ova_df1.columns:
                if not ova_df1[name].isna().all():
                    df1_amount_col = name
                    break
        df2_amount_col = ""
        for name in amount_col_names:
            if name in ova_df2.columns:
                if not ova_df2[name].isna().all():
                    df2_amount_col = name
                    break
        
        ova_value = ova_df1[df1_amount_col].sum() + ova_df2[df2_amount_col].sum()
        print(f"{file_output_name} OVA VOLUME : {ova_volume}")
        print(f"{file_output_name} OVA VALUE : {ova_value}")

        with pd.ExcelWriter(
            recons_file, engine="openpyxl", mode="w"
        ) as writer:  # specify new file name to write to
            ova_df1.to_excel(
                writer, sheet_name="Sheet1", index=False
            )  # save original data into first sheet of new file
    except ValueError:            
        ova_files = tuple(item for item in ova_files if item is not None)
        ova_df1 = pd.read_excel(ova_files[0],skiprows=num_lines_of_header[0])
        ova_volume = len(ova_df1)

        if ova_files[0] is not None:
            recons_file = (f"{ova_files[0][:-5]} - Recons.xlsx")
        else:
            recons_file = ""

        amount_col = ""
        for name in amount_col_names:
            if name in ova_df1.columns:
                if not ova_df1[name].isna().all():
                    amount_col = name
                    break
        
        ova_value = ova_df1[amount_col].sum()
        print(f"{file_output_name} OVA VOLUME : {ova_volume}")
        print(f"{file_output_name} OVA VALUE : {ova_value}")

        with pd.ExcelWriter(
            recons_file, engine="openpyxl", mode="w"
        ) as writer:  # specify new file name to write to
            ova_df1.to_excel(
                writer, sheet_name="Sheet1", index=False
            )  # save original data into first sheet of new file
    return recons_file

def get_write_double_int_val(
        int_files: Tuple[str | None, str | None],
        num_lines_of_header: Tuple[int, int],
        alt_recons_name: str,
        file_output_name: str,
        recons_file: str
        ):
    try:
        int_df1 = pd.read_excel(int_files[0],skiprows=num_lines_of_header[0])
        int_df2 = pd.read_excel(int_files[1],skiprows=num_lines_of_header[1])
        int_df2 = int_df2.loc[int_df2["Status"]== "CONFIRMED"]

        int_volume = len(int_df1) + len(int_df2)

        df1_amount_col = ""
        for name in amount_col_names:
            if name in int_df1.columns:
                if not int_df1[name].isna().all():
                    df1_amount_col = name
                    break

        df2_amount_col = ""
        for name in amount_col_names:
            if name in int_df2.columns:
                if not int_df2[name].isna().all():
                    df2_amount_col = name
                    break
        
        int_value = int_df1[df1_amount_col].sum() + int_df2[df2_amount_col].sum()
        print(f"{file_output_name} INT VOLUME : {int_volume}")
        print(f"{file_output_name} INT VALUE : {int_value}")
        dup, dup_val = find_duplicates(int_df1)
        write_duplicate_data(amount_col_name=df1_amount_col,df=dup,value=dup_val,file_name=recons_file)
        dup2, dup_val = find_duplicates(int_df2)
        write_duplicate_data(amount_col_name=df2_amount_col,df=dup2,value=dup_val,file_name=recons_file,last_row_index=len(dup)+3)
        return int_volume,int_value
    except ValueError:
        try:            
            int_files = tuple(item for item in int_files if item is not None)
            if int_files[0] == f"MPGS KC{yesterday}.xlsx" :
                int_df1 = pd.read_excel(int_files[0],skiprows=num_lines_of_header[0])
                int_volume = len(int_df1)

                amount_col = ""
                for name in amount_col_names:
                    if name in int_df1.columns:
                        if not int_df1[name].isna().all():
                            amount_col = name
                            break
                
                int_value = int_df1[amount_col].sum()
                print(f"{file_output_name} INT VOLUME : {int_volume}")
                print(f"{file_output_name} INT VALUE : {int_value}")
                dup, dup_val = find_duplicates(int_df1)
                write_duplicate_data(amount_col_name=amount_col,df=dup,value=dup_val,file_name=recons_file)
                return int_volume, int_value
            elif int_files[0] == f'MPGS_trn{yesterday}.xlsx':
                int_df1 = pd.read_excel(int_files[0],skiprows=num_lines_of_header[0])
                int_df1 = int_df1.loc[int_df1["Status"]== "CONFIRMED"]
                int_volume = len(int_df1)

                amount_col = ""
                for name in amount_col_names:
                    if name in int_df1.columns:
                        if not int_df1[name].isna().all():
                            amount_col = name
                            break
                
                int_value = int_df1[amount_col].sum()
                print(f"{file_output_name} INT VOLUME : {int_volume}")
                print(f"{file_output_name} INT VALUE : {int_value}")
                dup, dup_val = find_duplicates(int_df1)
                write_duplicate_data(amount_col_name=amount_col,df=dup,value=dup_val,file_name=recons_file)
                return int_volume, int_value           
        except:
            print("MPGS INT FILE NOT FOUND")




def find_duplicates(int_df: pd.DataFrame):
    trans_id_col = ""
    for name in tx_id_col_names:
        if name in int_df.columns:
            trans_id_col = name
            break

    amount_col = ""
    for name in amount_col_names:
        if name in int_df.columns:
            if not int_df[name].isna().all():
                amount_col = name
                break


    if trans_id_col == "":
        raise Exception("No transaction id column found")

    unique_tx = {}
    duplicates_tx = []

    duplicate_value: float = 0

    for index, row in int_df.iterrows():
        tx_id = row[trans_id_col]
        if tx_id in unique_tx:  # duplicate
            duplicates_tx.append(row)
            duplicate_value += row[amount_col]
        else:
            unique_tx[tx_id] = index

    return pd.DataFrame(duplicates_tx), duplicate_value


def write_duplicate_data(
    amount_col_name: str, df: pd.DataFrame, value: float, file_name: str, last_row_index=None
):
    if df.empty:
        return
    number_of_duplicates = len(df)
    amount_column = df.columns.get_loc(amount_col_name)
    if last_row_index is None:
        last_row_index = df.shape[0]
    empty_rows = pd.DataFrame({col: [None] for col in df.columns}, index=range(last_row_index, last_row_index + 3))
    empty_rows.iat[-2, amount_column] = value
    empty_rows.iat[-1, amount_column] = number_of_duplicates

    df = pd.concat([df, empty_rows], ignore_index=True)

    mode = "a" if os.path.exists(file_name) else "w"
    with pd.ExcelWriter(file_name, engine="openpyxl", mode=mode) as writer:
        sheet_name = "Duplicates"
        df.to_excel(writer, sheet_name=sheet_name, index=False)


def write_missing_ova_data(
    amount_col_name: str, df: pd.DataFrame, value: float, file_name: str, last_row_index=None
):
    if df.empty:
        return
    # If last_row_index is not specified, start from the end of the existing data
    if last_row_index is None:
        last_row_index = df.shape[0]
    number_of_tx = len(df)
    empty_rows = pd.DataFrame({col: [None] for col in df.columns}, index=range(last_row_index, last_row_index + 3))
    amount_column = df.columns.get_loc(amount_col_name)
    # Concatenate the empty rows with the original DataFrame
    df = pd.concat([df, empty_rows], ignore_index=True)
    # Update the number of tx and value in the new rows
    df.iat[last_row_index + 1, amount_column] = value
    df.iat[last_row_index + 2, amount_column] = number_of_tx
    with pd.ExcelWriter(file_name, engine="openpyxl", mode="a") as writer:
        sheet_name = "Missing OVA Transactions"
        df.to_excel(writer, sheet_name=sheet_name, index=False)


def write_missing_int_data(
    amount_col_name: str, df: pd.DataFrame, value: float, file_name: str, last_row_index=None
):
    if df.empty:
        return
    # If last_row_index is not specified, start from the end of the existing data
    if last_row_index is None:
        last_row_index = df.shape[0]
    number_of_tx = len(df)
    empty_rows = pd.DataFrame({col: [None] for col in df.columns}, index=range(last_row_index, last_row_index + 3))
    amount_column = df.columns.get_loc(amount_col_name)
    # Concatenate the empty rows with the original DataFrame
    df = pd.concat([df, empty_rows], ignore_index=True)
    # Update the number of tx and value in the new rows
    df.iat[last_row_index + 1, amount_column] = value
    df.iat[last_row_index + 2, amount_column] = number_of_tx
    with pd.ExcelWriter(file_name, engine="openpyxl", mode="a") as writer:
        sheet_name = "Missing INT Transactions"
        df.to_excel(writer, sheet_name=sheet_name, index=False)


def run_recons(
    file_names: Tuple[str | None, str | None],
    num_lines_of_header: Tuple[int, int],
    alt_recons_name: str,
    file_output_name: str,
    *,
    mb_service_name: str | None = None,
    mb_creditDebit_flag: str | None = None,
    mb_status_flag: str | None = None
):
    service_name_header= ""
    creditDebit_header= ""
    service_name_headers = ["ServiceName","serviceName"]
    creditDebit_headers=["creditDebitFlag","CreditDebitFlag","DEBITCREDIT"]
    # -------------------- OVA -------------------
    ova_file_name = file_names[0]  # name of the ova file
    int_file_name = file_names[1]  # name of the integrator file
    ova_header_lines = num_lines_of_header[0]
    int_header_lines = num_lines_of_header[1]
    recons_file = (
        f"{alt_recons_name} - Recons.xlsx"
        if ova_file_name is None
        else f"{ova_file_name[:-5]} - Recons.xlsx"
    )
    ova_file_df: pd.DataFrame | None = None
    int_file_df: pd.DataFrame | None = None

    if ova_file_name is not None:

        # put the data into a dataframe
        ova_file_df = pd.read_excel(ova_file_name,skiprows=ova_header_lines)
        ova_id_name = get_ova_id_col(ova_file_df)
        for name in creditDebit_headers:
            if name in ova_file_df.columns:
                    ova_file_df = ova_file_df[ova_file_df[name]== "C"]
                    break 
        ova_volume = len(ova_file_df)
        print(
            f"{file_output_name}_OVA_Volume: {ova_volume}"
        )  # file_output_name is the name that shows for each channel as the script runs
        amount_col = ""
        for name in amount_col_names:
            if name in ova_file_df.columns:
                if not ova_file_df[name].isna().all():
                    amount_col = name
                    break  # check which of the formats the amount column is written in

        ova_value = ova_file_df[amount_col].sum()

        print(f"{file_output_name} OVA_VALUE : {ova_value}")
        with pd.ExcelWriter(
            recons_file, engine="openpyxl", mode="w"
        ) as writer:  # specify new file name to write to
            ova_file_df.to_excel(
                writer, sheet_name="Sheet1", index=False
            )  # save original data into first sheet of new file

    

    # ----------------------- INTEGRATOR/ DUPLICATES --------------------
    if int_file_name is not None:
        int_file_df = pd.read_excel(int_file_name,skiprows=int_header_lines)
        int_id_name = get_int_id_col(int_file_df)
        for name in service_name_headers:
            if name in int_file_df.columns:
                if not int_file_df[name].isna().all():
                    service_name_header = name
                    break 
        for name in creditDebit_headers:
            if name in int_file_df.columns:
                if not int_file_df[name].isna().all():
                    creditDebit_header = name
                    break     
        if mb_service_name is not None and mb_creditDebit_flag is not None and mb_status_flag is None:
            int_file_df = int_file_df.loc[(int_file_df[service_name_header] == mb_service_name) & (int_file_df[creditDebit_header] == mb_creditDebit_flag)]
        elif mb_service_name is not None and mb_creditDebit_flag is None and mb_status_flag is None:
            int_file_df = int_file_df.loc[int_file_df[service_name_header] == mb_service_name]
        if mb_status_flag is not None and mb_service_name is None and mb_creditDebit_flag is None:
            int_file_df = int_file_df.loc[int_file_df["Status"] == mb_status_flag]
        if int_file_df.empty:
            print(f"Problem in {ova_file_name}. Check conditional headers")
            return
        int_volume = len(int_file_df)
        print(f"{file_output_name}_INT_Volume: {str(int_volume)}")
        amount_col = ""
        for name in amount_col_names:
            if name in int_file_df.columns:
                if not int_file_df[name].isna().all():
                    amount_col = name
                    break
        int_value = int_file_df[amount_col].sum()
        print(f"{file_output_name} INT_VALUE : {int_value}")
        dup, dup_val = find_duplicates(int_file_df)
        write_duplicate_data(
            amount_col_name=amount_col, df=dup, value=dup_val, file_name=recons_file
        )

    # ---------------------- MISSING TRANSACTIONS -------------------------
    if ova_file_df is not None and int_file_df is not None:
        ova_id_name =get_ova_id_col(ova_file_df)
        int_id_name = get_int_id_col(int_file_df)
        ova_file_df = ova_file_df.apply(lambda x: x.astype(str).str.lower())
        int_file_df = int_file_df.apply(lambda x: x.astype(str).str.lower())
        missing_int_tx = get_missing_tx(
            x=ova_file_df[ova_id_name].astype("string"),
            y=int_file_df[int_id_name].astype("string"),
        ).values
        missing_ova_tx = get_missing_tx(
            x=int_file_df[int_id_name].astype("string"),
            y=ova_file_df[ova_id_name].astype("string"),
        ).values
        int_amount_col = ""
        for name in amount_col_names:
            if name in int_file_df.columns:
                if not int_file_df[name].isna().all():
                    int_amount_col = name
                    break

        ova_amount_col = ""
        for name in amount_col_names:
            if name in ova_file_df.columns:
                if not ova_file_df[name].isna().all():
                    ova_amount_col = name
                    break
        ova_file_df[ova_amount_col] = ova_file_df[ova_amount_col].astype("float")
        int_file_df[int_amount_col]  = int_file_df[int_amount_col].astype("float")
        missing_ova_amount_name = ova_amount_col
        missing_int_amount_name = int_amount_col

        missing_ova_data = int_file_df[int_file_df[int_id_name].isin(missing_ova_tx)]
        missing_ova_value = missing_ova_data[missing_int_amount_name].sum()

        missing_int_data = ova_file_df[ova_file_df[ova_id_name].isin(missing_int_tx)]
        missing_int_value = missing_int_data[missing_ova_amount_name].sum()

        write_missing_ova_data(
            amount_col_name=int_amount_col,
            df=missing_ova_data,
            file_name=recons_file,
            value=missing_ova_value,
        )
        write_missing_int_data(
            amount_col_name=ova_amount_col,
            df=missing_int_data,
            file_name=recons_file,
            value=missing_int_value,
        )

def gip_custom(
ova_files: Tuple[str | None, str | None],
int_file: str,
num_lines_of_header: Tuple[int, int],
alt_recons_name: str,
file_output_name: str,
):
# ------------------------ OVA ------------------------
    recons_file = (
        f"{alt_recons_name} - Recons.xlsx"
        if ova_files[0] is None
        else f"{ova_files[0][:-5]} - Recons.xlsx"
)
    ova_files = tuple(item for item in ova_files if item is not None)
    if len(ova_files) == 2:
        ova_df1 = pd.read_excel(ova_files[0])
        ova_df2 = pd.read_excel(ova_files[1])

        ova_volume = len(ova_df1) + len(ova_df2)

        amount_col = ""
        for name in amount_col_names:
            if name in ova_df1.columns:
                if not ova_df1[name].isna().all():
                    amount_col = name
                    break
        ova_value = ova_df1[amount_col].sum() + ova_df2[amount_col].sum()
        print(f"GIP OVA VOLUME : {ova_volume}")
        print(f"GIP OVA VALUE : {ova_value}")

        with pd.ExcelWriter(
            recons_file, engine="openpyxl", mode="w"
        ) as writer:  # specify new file name to write to
            ova_df1.to_excel(
                writer, sheet_name="Sheet1", index=False
            )  # save original data into first sheet of new file
# ---------------------------- INT -------------------------
    int_df = pd.read_excel(int_file)
    int_volume = len(int_df)
    amount_col = ""
    for name in amount_col_names:
        if name in int_df.columns:
            if not int_df[name].isna().all():
                amount_col = name
                break
    
    int_value = int_df[amount_col].sum()
    print(f"GIP INT VOLUME {int_volume}")
    print(f"GIP INT VALUE {int_value}")
    dup, dup_val = find_duplicates(int_df)
    write_duplicate_data(amount_col_name=amount_col,df=dup,value=dup_val,file_name=recons_file)


def bb_mig_custom(
    ova_files: Tuple[str | None, str | None],
    int_file: str,
    num_lines_of_header: Tuple[int, int],
    alt_recons_name: str,
    file_output_name: str,    
):
    # ------------------------ OVA ------------------------
    recons_file= get_write_double_ova_val(ova_files=ova_files,num_lines_of_header=num_lines_of_header,file_output_name="BB MIG",alt_recons_name="BB MIG")
    

# ---------------------------- INT -------------------------
    int_df = pd.read_excel(int_file)
    int_df = int_df.loc[int_df["Status"]== "CONFIRMED"]
    int_volume = len(int_df)

    amount_col = ""
    for name in amount_col_names:
        if name in int_df.columns:
            if not int_df[name].isna().all():
                amount_col = name
                break
    int_value = int_df[amount_col].sum()
    print(f"BB MIG INT VOLUME {int_volume}")
    print(f"BB MIG INT VALUE {int_value}")
    dup, dup_val = find_duplicates(int_df)
    write_duplicate_data(amount_col_name=amount_col,df=dup,value=dup_val,file_name=recons_file)

def mpgs_custom(
    ova_files: Tuple[str | None, str | None],
    int_files: Tuple[str|None, str|None],
    num_lines_of_header: Tuple[int, int],
    alt_recons_name: str,
    file_output_name: str, 
):
    recons_file = get_write_double_ova_val(ova_files=ova_files,num_lines_of_header=num_lines_of_header,file_output_name=file_output_name,alt_recons_name=alt_recons_name)
    get_write_double_int_val(int_files=int_files,num_lines_of_header=num_lines_of_header,alt_recons_name=alt_recons_name,file_output_name=file_output_name,recons_file=recons_file) # type: ignore

    
def get_missing_tx(x: pd.Series, y: pd.Series) -> pd.Series:
    """
    Returns the elements in Series 'x' that are not present in Series 'y'.

    Parameters:
    ----------
    x (pd.Series): A pandas Series containing elements to be checked for presence in 'y'.
    y (pd.Series): A pandas Series containing elements to be checked against for presence.

    Returns:
    -------
    pd.Series: A pandas Series containing elements from 'x' that are missing in 'y'.

    Note:
    ----
    - This function performs a check to find elements in Series 'x' that are not present in Series 'y'.
    - The resulting Series contains only the elements that are missing in 'y' while maintaining the original order from 'x'.
    """
    x= remove_leading_zeros(x)
    y=remove_leading_zeros(y)
    missing = ~x.isin(y)
    return x[missing]

def remove_leading_zeros(series):
    # Remove leading zeros from the series
    return series.str.replace(r'^0+', '', regex=True)
if __name__ == "__main__":
    prompt = input("Recons for yesterday? (Y/N) :")
    if prompt.upper() == "Y":
        date_ = date.today() - timedelta(1)
    else:
        print("Please enter the date...")
        day = int(input("Recons Day (1-31): "))
        month = int(input("Recons Month (1-12): "))
        year = int(input("Recons Year (eg 2023): "))
        date_ = datetime(year=year, month=month, day=day)

    yesterday = f"_{get_date(date_, format='normal')}"
    print(yesterday)

    recons_yesterday = f"{get_date(date_, format='recons')} 00:00:00"
    print(recons_yesterday)

    current_month = date_.strftime("%b").upper()  # JAN, FEB
    print(current_month)

    GIPdate = get_date(date=date_ + timedelta(1), format="gip")
    print(GIPdate)
#TODO: find a way to make code select the service name and credit debit flag  on its own
    # ova_volumes = []
    # run_recons(
    #     (check_for_file(f"MIGS 01{yesterday}.xlsx"), check_for_file(f"MIGS 01 Metabase{yesterday}.xlsx")),
    #     num_lines_of_header=(3, 0),
    #     alt_recons_name=f"MIGS 01{yesterday}",
    #     file_output_name="MIGS_01",   
    # )

    # run_recons(
    #     (check_for_file(f"MTN Prompt{yesterday}.xlsx"), check_for_file(f"Metabase{yesterday}.xlsx")),
    #     num_lines_of_header=(0, 0),
    #     mb_service_name="MTN OVA",
    #     mb_creditDebit_flag="C",
    #     alt_recons_name=f"MTN Prompt{yesterday}",
    #     file_output_name="MTN Prompt",
    # )
    
    # run_recons(
    #     (check_for_file(f"MTN Cashout{yesterday}.xlsx"),check_for_file(f"Metabase{yesterday}.xlsx")),
    #     num_lines_of_header= (0,0),
    #     alt_recons_name="MTN Cashout",
    #     file_output_name="MTN_PORTAL",
    #     mb_service_name="MTN OVA", #TODO: USE MADAPI FOR CURRENT DATA
    #     mb_creditDebit_flag="D"
    # )
    # run_recons(
    #     (check_for_file(f"AirtelTigo Cashout{yesterday}.xlsx"),check_for_file(f"Metabase{yesterday}.xlsx")),
    #     num_lines_of_header= (4,0),
    #     alt_recons_name="AirtelTigo Cashout",
    #     file_output_name="AIRTEL_CASHOUT",
    #     mb_service_name="AirtelMoney_Slydepay",
    #     mb_creditDebit_flag="D"
    # )
    # run_recons(
    #     (check_for_file(f"Vodafone Cashin{yesterday}.xlsx"),check_for_file(f"Metabase{yesterday}.xlsx")),
    #     num_lines_of_header= (5,0),
    #     alt_recons_name="Vodafone Cashin",
    #     file_output_name="VODA CASHIN",
    #     mb_service_name="Vodafone Cash",
    #     mb_creditDebit_flag="C"
    #     #TODO: GET CODE TO EXTRACT IDS FROM OTHER COLUMN AND DO YOUR THING
    # )
    # run_recons(
    #     (check_for_file(f"Vodafone Cashout{yesterday}.xlsx"),check_for_file(f"Metabase{yesterday}.xlsx")),
    #     num_lines_of_header= (5,0),
    #     alt_recons_name="Vodafone Cashout",
    #     file_output_name="VODA CASHOUT",
    #     mb_creditDebit_flag="D",
    #     mb_service_name="Vodafone Cash"
    # )
    # run_recons(
    #     (check_for_file(f"KR MTN Credit{yesterday}.xlsx"),check_for_file(f"KR MTN Disb_mBase{yesterday}.xlsx")),
    #     num_lines_of_header= (0,0),
    #     alt_recons_name="KR MTN Credit",
    #     file_output_name="MTN_KR_Credit",
    # )
    # run_recons(
    #     (check_for_file(f"KR MTN Debit{yesterday}.xlsx"),check_for_file(f"KR MTN Coll_mBase{yesterday}.xlsx")),
    #     num_lines_of_header= (0,0),
    #     alt_recons_name="KR MTN Debit",
    #     file_output_name="MTN_KR_Debit",
    # )
    # run_recons(
    #     (check_for_file(f"KR Vodafone Cashin{yesterday}.xlsx"),check_for_file(f"KR Vodafone Coll_mBase{yesterday}.xlsx")),
    #     num_lines_of_header= (5,0),
    #     alt_recons_name="KR VODA Cashin",
    #     file_output_name="Voda KR Cashin",
    #     mb_status_flag="CONFIRMED"
    # )

    # run_recons(
    #     (check_for_file(f"KR Vodafone Cashout{yesterday}.xlsx"),check_for_file(f"KR Vodafone Disb_mBase{yesterday}.xlsx")),
    #     num_lines_of_header= (5,0),
    #     alt_recons_name="KR Voda Cashout",
    #     file_output_name="Voda KR Cashout",
    #     mb_status_flag="CONFIRMED"
    # )
    run_recons(
        (check_for_file(f"Stanbic FI Credit{yesterday}.xlsx"),check_for_file(f"Stanbic FI Credit Metabase{yesterday}.xlsx")),
        num_lines_of_header= (0,0),
        alt_recons_name="Stanbic FI Credit",
        file_output_name="Stanbic FI CREDIT",
        mb_status_flag="CONFIRMED"
    )
#     gip_custom(
#         ova_files=(check_for_file(f"slydepay_sending_{GIPdate}'.xlsx"),check_for_file(f"slydepay_sendingGhlink_{GIPdate}.xlsx")),
#         num_lines_of_header=(0,0),
#         alt_recons_name="slydepay_sending_",
#         file_output_name="GIP",
#         int_file=f'GIP Metabase{yesterday}.xlsx'
#    )
    # bb_mig_custom(
    #     ova_files=(check_for_file(f'MIGS 08{yesterday}.xlsx'),check_for_file(f'MIGS 09{yesterday}.xlsx')),
    #     num_lines_of_header=(3,3),
    #     alt_recons_name=f'MIGS 08{yesterday}.xlsx',
    #     file_output_name="BB MIG",
    #     int_file=f'MiGS_trn{yesterday}.xlsx'
    # )
#---------------------------
    # mpgs_custom(
    #     ova_files=(check_for_file(f'MPGS{yesterday}.xlsx'),check_for_file(f'Quipu{yesterday}.xlsx')),
    #     num_lines_of_header=(0,0),
    #     alt_recons_name=f'MPGS{yesterday}',
    #     file_output_name="MPGS",
    #     int_files=(check_for_file(f'MPGS KC{yesterday}.xlsx'),check_for_file(f'MPGS_trn{yesterday}.xlsx')),
    # )

