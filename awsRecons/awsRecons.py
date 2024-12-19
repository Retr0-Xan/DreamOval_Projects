import pandas as pd
import boto3

from typing import Literal, Tuple
from openpyxl import load_workbook
from datetime import datetime
from datetime import date
from datetime import timedelta

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
    "REFERENCE_NUMBER",
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
    "Total",
]

OVA_CHANNELS = {
        "KB_MOMO_MTN_Collection":"MTN-GH-Collections",
        "KB_MOMO_MTN_Disbursement":"MTN-GH-Disbursements",
        "KB_MOMO_VODAFONE_Collection":"Vodafone-GH-Collections",
        "KB_MOMO_VODAFONE_Disbursement":"Vodafone-GH-Disbursements",
        "NGENIUS":"Card-GH-NGENIUS",
        "KB_CARD_GT_Transactions":"Card-GH-GTMPGS",
        "SecurePay_Collections":"SecurePay-GH-Collections",
        "SecurePay_Disbursements":"SecurePay-GH-Disbursements",
    }

def get_date(
    date: date, format: Literal["gip"] or Literal["normal"] or Literal["recons"]
):
    """
    Formats a given date according to the specified format.

    Parameters:
    date (date): A datetime.date object representing the date to be formatted.
    format (Literal["gip"] or Literal["normal"] or Literal["recons"]): The format to use for formatting the date.
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
        return date.strftime("%Y-%m-%d")  # Output data in recons: 2023-01-01
    else:
        return date.strftime("%d %b_%y")  # Output date in format: 1_Jan_23


########################### update recons sheet logic here ######################################


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
    amount_col_name: str,
    df: pd.DataFrame,
    value: float,
    file_name: str,
    last_row_index=None,
):
    if df.empty:
        return
    number_of_duplicates = len(df)
    amount_column = df.columns.get_loc(amount_col_name)
    if last_row_index is None:
        last_row_index = df.shape[0]
    empty_rows = pd.DataFrame(
        {col: [None] for col in df.columns},
        index=range(last_row_index, last_row_index + 3),
    )
    empty_rows.iat[-2, amount_column] = value
    empty_rows.iat[-1, amount_column] = number_of_duplicates

    df = pd.concat([df, empty_rows], ignore_index=True)

    mode = "a" if os.path.exists(file_name) else "w"
    with pd.ExcelWriter(file_name, engine="openpyxl", mode=mode) as writer:
        sheet_name = "Duplicates"
        df.to_excel(writer, sheet_name=sheet_name, index=False)

def write_missing_ova_data(
    amount_col_name: str,
    df: pd.DataFrame,
    value: float,
    file_name: str,
    last_row_index=None,
):
    if df.empty:
        return
    # If last_row_index is not specified, start from the end of the existing data
    if last_row_index is None:
        last_row_index = df.shape[0]
    number_of_tx = len(df)
    empty_rows = pd.DataFrame(
        {col: [None] for col in df.columns},
        index=range(last_row_index, last_row_index + 3),
    )
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
    amount_col_name: str,
    df: pd.DataFrame,
    value: float,
    file_name: str,
    last_row_index=None,
):
    if df.empty:
        return
    # If last_row_index is not specified, start from the end of the existing data
    if last_row_index is None:
        last_row_index = df.shape[0]
    number_of_tx = len(df)
    empty_rows = pd.DataFrame(
        {col: [None] for col in df.columns},
        index=range(last_row_index, last_row_index + 3),
    )
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
    ova_file: str or None,
    int_df: pd.DataFrame or None,
    ova_file_key:str or None,
    alt_recons_name: str,
    file_output_name: str,
    ova_id: str,
    int_id: str,
    alt_ova_id: str,
    alt_int_id: str or None = None,
    *,
    mb_service_name: str or None = None,
    mb_creditDebit_flag: str or None = None,
    mb_status_flag: str or None = None,
    ova_status_flag: str or None = None,
    ova_status_col: str or None = None,
    list_index: int,
):
    if ova_file is None and int_df is None:
        return
    service_name_header = ""
    creditDebit_header = ""
    service_name_headers = ["ServiceName", "serviceName", "Service Name"]
    creditDebit_headers = [
        "creditDebitFlag",
        "CreditDebitFlag",
        "DEBITCREDIT",
        "Credit Debit Flag",
    ]
    # -------------------- OVA -------------------
    ova_file_name = ova_file  # name of the ova file
    recons_file = (
        f"{alt_recons_name} - Recons.xlsx"
        if ova_file_name is None
        else f"{ova_file_name[:-5]} - Recons.xlsx"
    )

    if ova_file_name is not None:
        # put the data into a dataframe
        ova_response = s3_client.get_object(Bucket=bucket_name, Key=ova_file_key)
        ova_file_df = pd.read_csv(ova_response['Body'])
        ova_id_name = ova_id
        if ova_status_flag is not None:
            ova_file_df = ova_file_df.loc[
                ova_file_df[ova_status_col] == ova_status_flag
            ]

        for name in creditDebit_headers:
            if name in ova_file_df.columns:
                ova_file_df = ova_file_df[ova_file_df[name] == "C"]
                break
        ova_volume = len(ova_file_df)
        ova_volumes[list_index] = ova_volume
        print(
            f"{file_output_name}_OVA_Volume: {ova_volume}"
        )  # file_output_name is the name that shows for each channel as the script runs
        amount_col = ""
        for name in amount_col_names:
            if name in ova_file_df.columns:
                if not ova_file_df[name].isna().all():
                    amount_col = name
                    break  # check which of the formats the amount column is written in

        ova_value = ova_file_df[amount_col].abs().sum()
        ova_values[list_index] = ova_value
        print(f"{file_output_name} OVA_VALUE : {ova_value}")
        with pd.ExcelWriter(
            recons_file, engine="openpyxl", mode="w"
        ) as writer:  # specify new file name to write to
            ova_file_df.to_excel(
                writer, sheet_name="Sheet1", index=False
            )  # save original data into first sheet of new file

    # ----------------------- INTEGRATOR/ DUPLICATES --------------------
    if int_df is not None:
        int_file_df = int_df
        int_id_name = int_id
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
        if (
            mb_service_name is not None
            and mb_creditDebit_flag is not None
            and mb_status_flag is None
        ):
            int_file_df = int_file_df.loc[
                (int_file_df[service_name_header] == mb_service_name)
                & (int_file_df[creditDebit_header] == mb_creditDebit_flag)
            ]
        elif (
            mb_service_name is not None
            and mb_creditDebit_flag is None
            and mb_status_flag is None
        ):
            int_file_df = int_file_df.loc[
                int_file_df[service_name_header] == mb_service_name
            ]
        if (
            mb_status_flag is not None
            and mb_service_name is None
            and mb_creditDebit_flag is None
        ):
            int_file_df = int_file_df.loc[int_file_df["Status"] == mb_status_flag]
        if int_file_df.empty:
            print(f"No transactions found for {file_output_name}. Confirm")
            return
        int_volume = len(int_file_df)
        int_volumes[list_index] = int_volume
        print(f"{file_output_name}_INT_Volume: {str(int_volume)}")
        amount_col = ""
        for name in amount_col_names:
            if name in int_file_df.columns:
                if not int_file_df[name].isna().all():
                    amount_col = name
                    break
        int_value = int_file_df[amount_col].abs().sum()
        int_values[list_index] = int_value

        print(f"{file_output_name} INT_VALUE : {int_value}")
        dup, dup_val = find_duplicates(int_file_df)
        print(f"Number of duplicates: {len(dup)}")
        print(f"Duplicates value: {dup_val}")
        dup_volumes[list_index] = len(dup)
        dup_values[list_index] = dup_val

        write_duplicate_data(
            amount_col_name=amount_col, df=dup, value=dup_val, file_name=recons_file
        )
    if ova_file_name == f"MPGS{yesterday}.xlsx":
        return
    # ---------------------- MISSING TRANSACTIONS -------------------------
    if ova_file_df is not None and int_file_df is not None:
        ova_id_name = ova_id
        int_id_name = int_id
        ova_file_df[ova_id_name] = ova_file_df[ova_id_name].astype(str)
        int_file_df[int_id_name] = int_file_df[int_id_name].astype(str)

        missing_int_tx = get_missing_tx(
            x=ova_file_df[ova_id_name].astype("string"),
            y=int_file_df[int_id_name].astype("string"),
            alt_x=ova_file_df[alt_ova_id].astype("string"),
            alt_y=int_file_df[alt_int_id].astype("string"),
        ).values

        missing_ova_tx = get_missing_tx(
            x=int_file_df[int_id_name].astype("string"),
            y=ova_file_df[ova_id_name].astype("string"),
            alt_x=int_file_df[alt_int_id].astype("string"),
            alt_y=ova_file_df[alt_ova_id].astype("string"),
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
        int_file_df[int_amount_col] = int_file_df[int_amount_col].astype("float")
        missing_ova_amount_name = ova_amount_col
        missing_int_amount_name = int_amount_col

        missing_ova_data = int_file_df[
            int_file_df[int_id_name].astype("string").isin(missing_ova_tx)
            | int_file_df[alt_int_id].astype("string").isin(missing_ova_tx)
        ]
        missing_ova_value = missing_ova_data[missing_int_amount_name].abs().sum()

        missing_int_data = ova_file_df[
            ova_file_df[ova_id_name].astype("string").isin(missing_int_tx)
            | ova_file_df[alt_ova_id].astype("string").isin(missing_int_tx)
        ]

        missing_int_value = missing_int_data[missing_ova_amount_name].abs().sum()

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

def check_for_file():
    pass


def  get_missing_tx(
    x: pd.Series,
    y: pd.Series,
    alt_x: pd.Series,
    alt_y: pd.Series,
) -> pd.Series:
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
    x = remove_leading_zeros(x)
    y = remove_leading_zeros(y)

    missing = (~x.astype(str).str.lower().isin(y.astype(str).str.lower())) & (
        ~alt_x.astype(str).str.lower().isin(alt_y.astype(str).str.lower())
    )
    x_missing = x[missing].combine_first(alt_x[missing])
    x_missing[(x_missing == "nan")] = alt_x[missing]
    return x_missing


def remove_leading_zeros(series):
    # Remove leading zeros from the series
    return series.str.replace(r"^0+", "", regex=True)


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
    ova_volumes = [0] * 17
    ova_values = [0] * 17
    int_volumes = [0] * 17
    int_values = [0] * 17
    dup_volumes = [0] * 17
    dup_values = [0.00] * 17
    list_index = 0

    
    # Create an S3 client
    s3_client = boto3.client("s3")
    # List all S3 buckets
    buckets = s3_client.list_buckets()
    # List of unwanted headers for MTN transactions
    bucket_name = 'all-kowri-datalake'



    metabase_file_key = f'KowriBusiness/KBPlatform-Transaction/year={year}/month={month}/day={day}/KBPlatform_transaction_{year}_{month}_{day}.csv'
    response = s3_client.get_object(Bucket=bucket_name, Key=metabase_file_key)
    main_df = pd.read_csv(response['Body'],delimiter='|')

    gip_df = main_df.loc[(main_df['accountId'] == '596ddebc4b026bff449f2c58')&(main_df['creditDebitFlag'] == 'C')&(main_df['status'] == 'CONFIRMED')]
    mtn_collections_df = main_df.loc[(main_df['accountId'] == '617a8b63200b660012110655')&(main_df['creditDebitFlag'] == 'C')&(main_df['status'] == 'CONFIRMED')]
    mtn_disbursements_df = main_df.loc[(main_df['accountId'] == '617a8b019575ed0012d8d23a')&(main_df['creditDebitFlag'] == 'C')&(main_df['status'] == 'CONFIRMED')]
    #negenius KC
    ngenius_KB_df = main_df.loc[((main_df['accountId'] == '66e86d5cdc02025d2a384359') | (main_df['accountId'] == '673ca05157c47c2404c43b05')) & (main_df['status'] == 'CONFIRMED')]
    securePay_collections_df = main_df.loc[(main_df['accountId'] == '667313dcaa0ad507c89caef6')&(main_df['status'] == 'CONFIRMED') & (main_df['transactionType'] == 'MOBILE_MONEY')]
    securePay_disbursements_df = main_df.loc[(main_df['accountId'] == '6673130c53a7000ceb99edbd')&(main_df['status'] == 'CONFIRMED') & (main_df['transactionType'] == 'MOBILE_MONEY')]
    telecel_collections_df = main_df.loc[(main_df['serviceName'] == 'Telecel Cash') & (main_df['status'] == 'CONFIRMED') & (main_df['transactionType'] != 'CHARGE') & (main_df['serviceProvider'] == 'Telecel Cash Kowri Payment OVA') & (main_df['creditDebitFlag'] == 'C')]
    telecel_disbursements_df = main_df.loc[(main_df['serviceName'] == 'Telecel Cash') & (main_df['status'] == 'CONFIRMED') & (main_df['transactionType'] != 'CHARGE') & (main_df['serviceProvider'] == 'Telecel Cash Kowri Send Money OVA') & (main_df['creditDebitFlag'] == 'C')]

    try:
        run_recons(
            f"KB_MOMO_MTN_Collection{yesterday}.csv",
            ova_file_key = f'KowriBusiness/MTN-GH-Collections/year={year}/month={month}/day={day}/KBPlatform_transaction_{year}_{month}_{day}.csv',
            num_lines_of_header=(0, 0),
            alt_recons_name=f"KR MTN Debit{yesterday}",
            file_output_name="MTN_KR_Debit",
            list_index=3,
            ova_id="External Transaction Id",
            int_id="integratorTransId",
            alt_ova_id="Id",
            alt_int_id="billerTransId",
        )
    except:
        ova_volumes[list_index] = 0
        int_volumes[list_index] = 0
        int_values[list_index] =0
        int_volumes[list_index] = 0
        run_recons(
            f"KB_MOMO_MTN_Collection{yesterday}.csv",
            ova_file_key = f'KowriBusiness/MTN-GH-Collections/year={year}/month={month}/day={day}/KBPlatform_transaction_{year}_{month}_{day}.csv',
            num_lines_of_header=(0, 0),
            alt_recons_name=f"KR MTN Debit{yesterday}",
            file_output_name="MTN_KR_Debit",
            list_index=3,
            ova_id="External id",
            int_id="integratorTransId",
            alt_ova_id="Id",
            alt_int_id="billerTransId",
        )


