import pandas as pd

ova_file = pd.read_excel("KR MTN Debit_30 Oct_24.xlsx")
int_file = pd.read_excel("KR MTN Coll_mBase_30 Oct_24.xlsx")

#compare id columns of both files
ova_file["External id"] = ova_file["External id"].astype(str)
int_file["IntegratorTransId"] = int_file["IntegratorTransId"].astype(str)

def remove_leading_zeros(series):
    # Remove leading zeros from the series
    return series.str.replace(r"^0+", "", regex=True)


def get_missing_tx(
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

missing_int_tx = get_missing_tx(
    x=ova_file["External id"].astype("string"),
    y=int_file["IntegratorTransId"].astype("string"),
    alt_x=ova_file["Id"].astype("string"),
    alt_y=int_file["BillerTransId"].astype("string"),
).values

print(missing_int_tx)


