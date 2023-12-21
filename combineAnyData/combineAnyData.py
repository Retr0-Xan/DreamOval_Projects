import pandas as pd
import os

#Function to lisy all the files in a directory
def list_files_recursive(directory):
    all_files = []
    for root, dirs, files in os.walk(directory):
        all_files.extend([os.path.join(root, file) for file in files])
    return all_files

#getting the location of the script
current_directory = os.getcwd()

#Listing all the files in this directory
files = list_files_recursive(f"{current_directory}/combineAnyData")

#Function to combine the data
def combine_data(output_name, ext):
    complete_file_df = pd.DataFrame()
    for file in files:
        print(file)
        if ext == "xlsx":
            if ext in file:
                complete_file_df = pd.concat([complete_file_df, pd.read_excel(file)])
        elif ext == "csv":
            if ext in file:
                print(file)
                complete_file_df = pd.concat([complete_file_df, pd.read_csv(file)])
    complete_file_df.to_csv(f"./combineAnyData/data/{output_name}.csv", index=False)


#Name of the final output file with the combined data
output_name = input("Please enter an output name for final file: ")

#The extension/datatype of the files that you are combining
extension = input(
    "What extension are your files in? (Please make sure all files are of the same type): "
)

combine_data(output_name=output_name, ext=extension)
