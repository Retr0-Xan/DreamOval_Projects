# version 1.1.0
import pandas as pd
from typing import Literal, Tuple

# from dateTime import yesterday, GIPdate, current_month, recons_yesterday
from datetime import datetime
from datetime import date
from datetime import timedelta


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

# def metabase_file_exists() -> bool:
#     dir_list = os.listdir()
#     # Check if metabase file is available
#     metabase_file_name = "Metabase" + yesterday + ".xlsx"
#     for file in dir_list:
#         if file == metabase_file_name:
#             return True
#     return False


def find_duplicates(int_df: pd.DataFrame):
    tx_id_col_names = ["integratorTransId", "IntegratorTransId"]
    amount_col_names = ["Amount", "amount"]

    trans_id_col = ""
    for name in tx_id_col_names:
        if name in int_df.columns:
            trans_id_col = name
            break

    amount_col = ""
    for name in amount_col_names:
        if name in int_df.columns:
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


def run_recons(
    file_names: Tuple[str | None, str | None],
    num_lines_of_header: Tuple[int, int],
    mb_service_name: str,
    mb_creditDebit_flag: str,
    alt_recons_name: str,
):
    ova_file_name = file_names[0]
    int_file_name = file_names[1]
    ova_header_lines = num_lines_of_header[0]
    int_header_lines = num_lines_of_header[1]
    recons_file = alt_recons_name if ova_file_name is None else f"{ova_file_name[:-5]} - Recons.xlsx"

    ova_file_df: pd.DataFrame | None = None
    int_file_df: pd.DataFrame | None = None
    if ova_file_name is not None:
        ova_file_df = pd.read_excel(ova_file_name)
        OVA_VOLUME1 = len(ova_file_df)
        print(f"MIGS_01_OVA_Volume: {str(OVA_VOLUME1)}")
        amount_col_names = ["Amount", "amount"]
        amount_col = ""
        for name in amount_col_names:
            if name in ova_file_df:
                amount_col = name
            break
        OVA_VALUE1 = ova_file_df[amount_col].sum()

        print(OVA_VALUE1)
        with pd.ExcelWriter(
            recons_file, engine="openpyxl", mode="w"
        ) as writer:  # specify new file name to write to
            ova_file_df.to_excel(
                writer, sheet_name="Sheet1", index=False
            )  # save original data into first sheet of new file

    if int_file_name is not None:
        int_file_df = pd.read_excel(int_file_name)
        int_file_df = int_file_df.loc[
            (int_file_df["ServiceName"] == mb_service_name)
            & (int_file_df["CreditDebitFlag"] == mb_creditDebit_flag)
        ]
        INT_VOLUME1 = len(int_file_df)
        print(f"MIGS_01_INT_Volume: {str(INT_VOLUME1)}")
        amount_col_names = ["Amount", "amount"]
        amount_col = ""
        for name in amount_col_names:
            if name in int_file_df:
                amount_col = name
            break
        INT_VALUE1 = int_file_df[amount_col].sum()
        print(INT_VALUE1)
        dup, dup_val = find_duplicates(int_file_df)

        with pd.ExcelWriter(recons_file, engine="openpyxl", mode="a") as writer:
            sheet_name = "Duplicates"

            dup.to_excel(writer, sheet_name=sheet_name, index=False)
            # TODO: Write dup_val to the sheet

    if ova_file_df and int_file_df:
        missing_ova_idx = get_missing_tx(
            x=ova_file_df["Id"].astype("string"),
            y=int_file_df["BillerTransId"].astype("string"),
        ).index

        missing_int_idx = get_missing_tx(
            x=int_file_df["BillerTransId"].astype("string"),
            y=ova_file_df["Id"].astype("string"),
        ).index

        missing_ova_tx = ova_file_df.iloc[missing_ova_idx]
        missing_int_tx = int_file_df.iloc[missing_int_idx]

        with pd.ExcelWriter(recons_file, engine="openpyxl", mode="a") as writer:
            sheet_name = "Missing OVA Transactions"
            missing_ova_tx.to_excel(writer, sheet_name=sheet_name, index=False)

        with pd.ExcelWriter(recons_file, engine="openpyxl", mode="a") as writer:
            sheet_name = "Missing Integrator Transactions"
            missing_int_tx.to_excel(writer, sheet_name=sheet_name, index=False)

            # TODO: Write missing ova anbd int transactions into sheet.


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
    missing = ~x.isin(y)
    return x[missing]


run_recons(
    ("MTN Prompt_13 Oct_23.xlsx", "Metabase_13 Oct_23.xlsx"),
    num_lines_of_header=(0, 0),
    mb_service_name="MTN OVA",
    mb_creditDebit_flag="C",
    alt_recons_name="MTN Prompt",
)
