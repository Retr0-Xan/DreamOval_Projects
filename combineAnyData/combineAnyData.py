import pandas as pd
import os


def list_files_recursive(directory):
    all_files = []
    for root, dirs, files in os.walk(directory):
        all_files.extend([os.path.join(root, file) for file in files])
    return all_files
current_directory = os.getcwd()

files =list_files_recursive(current_directory)

def combine_data(output_name, ext):
       complete_file_df = pd.DataFrame()
       for file in files:
            print(file)
            if ext in file:
                complete_file_df =pd.concat([complete_file_df,pd.read_excel(file)])
       complete_file_df.to_excel(f"./combineAnyData/data/{output_name}.{ext}",index=False)


output_name = input("Please enter an output name for final file: ")
extension = input("What extension are your files in? (Please make sure all files are of the same type): ")

combine_data(output_name=output_name,ext=extension)
