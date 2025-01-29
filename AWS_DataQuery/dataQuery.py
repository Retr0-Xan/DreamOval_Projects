import boto3
import pandas as pd
import os
import customtkinter as ctk
import tkinter as tk
from tkinter import Toplevel
from tkinter import ttk
from tkcalendar import Calendar
from PIL import Image, ImageTk
import sys
import platform
import threading
from datetime import datetime, timedelta
import subprocess
from tkinter import messagebox



def validate_credentials(access_key, secret_key):
    try:
        # Create a session with the credentials
        session = boto3.Session(
            aws_access_key_id=access_key,
            aws_secret_access_key=secret_key
        )
        # Use S3 client to validate the credentials
        s3 = session.client('s3')
        s3.list_buckets()  # Try listing buckets to test the credentials
        return session  # Return the session if valid
    except Exception as e:
        print(e)
        return None
    

def login_screen():
    def login():
        access_key = access_key_entry.get()
        secret_key = secret_key_entry.get()
        
        if not access_key or not secret_key:
            messagebox.showerror("Error", "Please enter both Access Key and Secret Key")
            return
        
        global aws_session
        aws_session = validate_credentials(access_key, secret_key)
        
        if aws_session:
            messagebox.showinfo("Success", "Login successful!")
            login_window.destroy()  # Close login screen
            main()
            
        else:
            messagebox.showerror("Error", "Invalid AWS Credentials")

    # Create login window
    login_window = ctk.CTk()
    login_window.title("AWS Login")
    login_window.geometry("400x300")

    # Access Key Input
    ctk.CTkLabel(login_window, text="AWS Access Key:").pack(pady=10)
    access_key_entry = ctk.CTkEntry(login_window)
    access_key_entry.pack(pady=10)

    # Secret Key Input
    ctk.CTkLabel(login_window, text="AWS Secret Key:").pack(pady=10)
    secret_key_entry = ctk.CTkEntry(login_window, show="*")
    secret_key_entry.pack(pady=10)

    # Login Button
    ctk.CTkButton(login_window, text="Login", command=login).pack(pady=20)

    login_window.mainloop()



def set_working_directory_to_script_location():
    if getattr(sys, "frozen", False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))

    os.chdir(script_dir)
    return script_dir

script_dir = set_working_directory_to_script_location()

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

stop_event = threading.Event()

def run_query(channel_var: tk.StringVar, from_date: ctk.CTkEntry, to_date: ctk.CTkEntry, name_entry: ctk.CTkEntry):
    # s3_client = boto3.client("s3")
    s3_client = aws_session.client("s3")

    unavailable_files = []
    bucket_name = 'all-kowri-datalake'
    current_directory = os.getcwd()
    channel_name = channel_var.get()

    start_date = datetime.strptime(from_date.get(), "%d-%m-%Y")
    end_date = datetime.strptime(to_date.get(), "%d-%m-%Y")
    output_file_name = name_entry.get()

    current_date = start_date
    file_name_locs = {
        "MTN-GH-Collections": "KB_MOMO_MTN_Collection",
        "MTN-GH-Disbursements": "KB_MOMO_MTN_Disbursement",
        "Vodafone-GH-Collections": "KB_MOMO_VODAFONE_Collection",
        "Vodafone-GH-Disbursements": "KB_MOMO_VODAFONE_Disbursement",
        "Card-GH-NGENIUS": "NGENIUS",
        "Card-GH-GTMPGS": "KB_CARD_GT_Transactions",
        "SecurePay-GH-Collections": "SecurePay_Collections",
        "SecurePay-GH-Disbursements": "SecurePay_Disbursements",
        "KBPlatform-MerchantOrder": "KBPlatform_merchantOrder",
        "KBPlatform-Transaction": "KBPlatform_transaction",
    }

    complete_file_df = pd.DataFrame()
    while current_date <= end_date:
        if stop_event.is_set():
            print("Download stopped.")
            break

        print("-------Gathering Data--------")
        current_year = str(current_date.year)
        current_month = str(current_date.month).zfill(2)
        current_day = str(current_date.day).zfill(2)
        file_key = f'KowriBusiness/{channel_name}/year={current_year}/month={current_month}/day={current_day}/{file_name_locs[channel_name]}_{current_year}_{current_month}_{current_day}.csv'
        file_name = f'{file_name_locs[channel_name]}_{current_year}_{current_month}_{current_day}.csv'
        print(file_key)

        try:
            response = s3_client.get_object(Bucket=bucket_name, Key=file_key)
            complete_file_df = pd.concat([complete_file_df, pd.read_csv(response['Body'])])
            print(f"Day {current_day}: Done...")
        except:
            try:
                complete_file_df = pd.concat([complete_file_df, pd.read_excel(response['Body'])])
                print(f"Day {current_day}: Done...")
            except:
                unavailable_files.append(file_name)
                print(f"Day {current_day}: Skipping...")
        current_date += timedelta(days=1)

        print(f'Unavailable Files: {unavailable_files}')


    if not stop_event.is_set():
        try:
            complete_file_df.to_csv(f"{current_directory}/data/{output_file_name}.csv", index=False)
        except Exception as e:
            print(e)
        folder_path = f"{current_directory}/data"
        if platform.system() == "Windows":
            os.startfile(folder_path)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", folder_path])
        else:
            subprocess.Popen(["xdg-open", folder_path])

def main():
    def open_date_picker(date_entry: ctk.CTkEntry, root: ctk.CTk):
        date_window = Toplevel(root)
        date_window.title("Select Date")
        date_window.geometry("350x350")

        cal = Calendar(date_window, selectmode="day", year=2024, month=11, day=6, date_pattern="dd-mm-yyyy")
        cal.pack(pady=20)

        def grab_date():
            selected_date = cal.get_date()
            date_entry.delete(0, ctk.END)
            date_entry.insert(0, selected_date)
            date_window.destroy()

        select_date_button = ctk.CTkButton(date_window, text="Select Date", command=grab_date, fg_color="green")
        select_date_button.pack(pady=10)

    root = ctk.CTk()
    root.title("DataQuery")
    root.option_add("*tearOff", False)
    root._set_appearance_mode("light")
    root.geometry("500x550")

    main_frame = ctk.CTkFrame(root, fg_color="white")
    main_frame.pack(fill="both", expand=True)

    logo_img = ctk.CTkImage(
        Image.open(resource_path(f"{script_dir}/assets/KowriLogo.png")),
        size=(200, 70),
    )

    label = ctk.CTkLabel(main_frame, image=logo_img, text="", fg_color="white")
    label.pack()

    channels = [
        "MTN-GH-Collections", "MTN-GH-Disbursements", "Vodafone-GH-Collections", "Vodafone-GH-Disbursements", "Card-GH-GTMPGS", "InstantPayment-GH",
        "Card-GH-NGENIUS", "SecurePay-GH-Collections", "SecurePay-GH-Disbursements", "KBPlatform-MerchantOrder", "KBPlatform-Transaction"
    ]
    channel_var = tk.StringVar(value=channels[0])

    channel_label = ctk.CTkLabel(main_frame, text="Select Channel:", text_color="green")
    channel_label.pack(pady=(20, 5))
    channel_dropdown = ctk.CTkOptionMenu(main_frame, variable=channel_var, values=channels, fg_color="white", button_color="green", dropdown_fg_color="white", text_color="black", dropdown_text_color="black", dropdown_hover_color="green")
    channel_dropdown.pack(pady=5)

    dates_frame = ctk.CTkFrame(main_frame, fg_color="white", bg_color="white")
    dates_frame.pack(pady=(20, 5))

    from_frame = ctk.CTkFrame(dates_frame, fg_color="white")
    from_frame.grid(row=0, column=0)
    from_label = ctk.CTkLabel(from_frame, text="From:", text_color="green")
    from_label.pack(pady=(20, 5))

    from_entry = ctk.CTkEntry(from_frame, fg_color="white", bg_color="white", text_color="black")
    from_entry.pack(pady=5, padx=20)

    date_button = ctk.CTkButton(from_frame, text="Select Date", command=lambda: open_date_picker(from_entry, root), text_color="white", fg_color="green")
    date_button.pack(pady=5)

    to_frame = ctk.CTkFrame(dates_frame, fg_color="white")
    to_frame.grid(row=0, column=1)
    to_label = ctk.CTkLabel(to_frame, text="To:", text_color="green")
    to_label.pack(pady=(20, 5))

    to_entry = ctk.CTkEntry(to_frame, fg_color="white", bg_color="white", text_color="black")
    to_entry.pack(pady=5)

    date_button = ctk.CTkButton(to_frame, text="Select Date", command=lambda: open_date_picker(to_entry, root), text_color="white", fg_color="green")
    date_button.pack(pady=5)

    outName_entry = ctk.CTkEntry(main_frame, fg_color="white", bg_color="white", text_color="black", placeholder_text="Enter Output Name", width=300)
    outName_entry.pack(pady=5)

    status_label = ctk.CTkLabel(main_frame, text="", text_color="green")
    status_label.pack(pady=(20, 5))

    def update_widgets(button: ctk.CTkButton, status):
        if status == "running":
            button.configure(fg_color="red", text="Stop", hover_color="dark red")
            status_label.configure(text="Gathering Data...", text_color="green")
        elif status == "done":
            button.configure(state="normal", fg_color="green", text="Run Query", hover_color="green")
            status_label.configure(text="Data Query Completed", text_color="green")
        elif status == "stopped":
            button.configure(state="normal", fg_color="green", text="Run Query", hover_color="green")
            status_label.configure(text="Query Stopped", text_color="red")

    def run_query_thread(channel_var, from_entry, to_entry, button, name_entry):
        stop_event.clear()
        update_widgets(button, "running")

        run_query(channel_var=channel_var, from_date=from_entry, to_date=to_entry, name_entry=name_entry)

        if not stop_event.is_set():
            update_widgets(button, "done")

    def stop_query(button):
        stop_event.set()
        update_widgets(button, "stopped")

    def check_data_validity(channel_var, from_entry, to_entry, button, name_entry):
        if not from_entry.get() or not to_entry.get() or not name_entry.get():
            status_label.configure(text="Please fill in all fields", text_color="red")
            return

        if button.cget("text") == "Stop":
            stop_query(button)
        else:
            threading.Thread(target=run_query_thread, args=(channel_var, from_entry, to_entry, button, name_entry)).start()

    submit_button = ctk.CTkButton(main_frame, text="Run Query", command=lambda: check_data_validity(channel_var, from_entry, to_entry, submit_button, outName_entry), text_color="white", fg_color="green")
    submit_button.pack(pady=(30, 10))

    style = ttk.Style(root)
    root.tk.call(
        "source",
        resource_path(f"{script_dir}/assets/Forest-ttk-theme-master/forest-light.tcl"),
    )

    style.theme_use("forest-light")
    root.mainloop()

if __name__ == "__main__":
    aws_session = None
    login_screen()
    # main()
