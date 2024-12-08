import pandas as pd

def split_sheets_to_csv(file_path, output_directory, base_filename):
    """
    Splits sheets from an Excel file and saves them as individual CSV files.
    
    Args:
        file_path (str): Path to the Excel file.
        output_directory (str): Directory to save the CSV files.
        base_filename (str): Base name for the output files.
        
    Returns:
        None
    """

    # Load the Excel file with multiple sheets
    sheets = pd.read_excel(file_path,sheet_name="mar 26")
    print(sheets)
    
    # Iterate through each sheet
    for sheet_name, data in sheets.items():
        print(sheet_name)
        # Convert sheet name to a formatted filename
        day = sheet_name.split()[-1]  # Extract day from sheet name
        month = sheet_name.split()[0].capitalize()  # Extract month and capitalize
        year = "24"  # Use '24' as the year
        
        formatted_filename = f"{base_filename}_{day} {month}_{year}.csv"
        
        # Save the sheet data to a CSV file
        output_path = f"{output_directory}\{formatted_filename}"
        data.to_csv(f"test {day}.csv", index=False)
        print(f"Saved: {output_path}")

# Example usage
if __name__ == "__main__":
    # Path to the Excel file
    file_path = "Telecel Disb 26_02-31_03.xlsx"
    
    # Output directory for CSV files
    output_directory = "C://Users//Mark//repos//DreamOval_Projects//AWS_DataQuery"
    
    # Base filename
    base_filename = "KR Telecel Cashout"
    
    # Split and save sheets
    split_sheets_to_csv(file_path, output_directory, base_filename)
