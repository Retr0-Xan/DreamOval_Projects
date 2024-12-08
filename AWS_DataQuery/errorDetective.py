import boto3
from botocore.exceptions import ClientError
from datetime import datetime,timedelta
import os


file_name_locs = {
    "MTN-GH-Collections":"KB_MOMO_MTN_Collection",
    "MTN-GH-Disbursements":"KB_MOMO_MTN_Disbursement",
    "Vodafone-GH-Collections":"KB_MOMO_VODAFONE_Collection",
    "Vodafone-GH-Disbursements":"KB_MOMO_VODAFONE_Disbursement",
    "Card-GH-NGENIUS":"NGENIUS",
    "Card-GH-GTMPGS":"KB_CARD_GT_Transactions",
    "SecurePay-GH-Collections":"SecurePay_Collections",
    "SecurePay-GH-Disbursements":"SecurePay_Disbursements",
    "KBPlatform-MerchantOrder":"KBPlatform_merchantOrder",
    "KBPlatform-Transaction":"KBPlatform_transaction",

}


def main():
    current_directory = os.path.dirname(os.path.abspath(__file__))
    detected_error_files = []

    # Create an S3 client
    s3_client = boto3.client("s3")
    # List all S3 buckets
    bucket_name = 'all-kowri-datalake'

    def rename_s3_file(bucket_name, old_file_key, new_file_key,alt_file_key):
        print("--------- Rename Attempt 1 -----------------------")
        try:
            # Copy the file to the new key
            s3_client.copy_object(Bucket=bucket_name, CopySource={'Bucket': bucket_name, 'Key': old_file_key}, Key=new_file_key)
            print(f'Copied {old_file_key} to {new_file_key}')

            # Delete the old file
            s3_client.delete_object(Bucket=bucket_name, Key=old_file_key)
            print(f'Deleted {old_file_key}')
            print(f"Renamed file: {old_file_key}")
            print("--------- Attempt 1 Successful -----------------------")
        except ClientError as e:
            # Check for NoSuchKey error
            if e.response['Error']['Code'] == 'NoSuchKey':
                try:
                    print("--------- Rename Attempt 2 -----------------------")
                    s3_client.copy_object(Bucket=bucket_name, CopySource={'Bucket': bucket_name, 'Key': alt_file_key}, Key=new_file_key)
                    print(f'Copied {alt_file_key} to {new_file_key}')

                    # Delete the old file
                    s3_client.delete_object(Bucket=bucket_name, Key=alt_file_key)
                    print(f'Deleted {alt_file_key}')
                    print(f"Renamed file: {alt_file_key}")
                    print("--------- Attempt 2 Successful -----------------------")
                    
                except:
                    print(f"Error renaming file: {alt_file_key}")
                    print(e)
                    # Define the file path for the text file
                    file_path = os.path.join(current_directory, "missing_data.txt")

                    # Append the result to the file
                    with open(file_path, "a") as file:
                        file.write( alt_file_key + "\n")  # Add a newline at the end of the result

                    print(f"Appended to {file_path}")
                    detected_error_files.append(alt_file_key)

    
    start_date = datetime.strptime("01-01-2024","%d-%m-%Y")
    end_date = datetime.strptime("18-11-2024","%d-%m-%Y")

    current_date = start_date
    channel_name = "Card-GH-NGENIUS"

    while current_date <= end_date:
        current_year = str(current_date.year)
        current_month = str(current_date.month).zfill(2)
        current_day = str(current_date.day).zfill(2)
        file_key = f'KowriBusiness/{channel_name}/year={current_year}/month={current_month}/day={current_day}/{file_name_locs[channel_name]}_{current_year}_{current_month}_{current_day}.csv'  # The key of the file you want to download
        alt_file_key = f'KowriBusiness/{channel_name}/year={current_year}/month={current_month}/day={current_day} /{file_name_locs[channel_name]}_{current_year}_{current_month}_{current_day}.csv'  # The key of the file you want to download
        error_file_key = f'KowriBusiness/{channel_name}/year={current_year}/month={current_month}/day={current_day} /{file_name_locs[channel_name]}_{current_year}_{current_month}_{current_day} .csv' 

        try:
            print(f"Checking availabilty - {channel_name}  {current_date}......... ")
            response = s3_client.get_object(Bucket=bucket_name, Key=file_key)
        except Exception as e:
            try:
                rename_s3_file(bucket_name,old_file_key=error_file_key,new_file_key=file_key,alt_file_key=alt_file_key)
            except Exception as e:
                print(f"Error renaming file: {error_file_key}")
                print(e)
                # Define the file path for the text file
                file_path = os.path.join(current_directory, "missing_data.txt")

                # Append the result to the file
                with open(file_path, "a") as file:
                    file.write( error_file_key + "\n")  # Add a newline at the end of the result

                print(f"Appended to {file_path}")
                detected_error_files.append(error_file_key)

        current_date += timedelta(days=1) 

    print(detected_error_files)


if __name__ == "__main__":
    main()



