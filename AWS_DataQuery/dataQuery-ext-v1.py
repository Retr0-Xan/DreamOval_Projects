import boto3
import pandas as pd
import os
import customtkinter as ctk
import tkinter as tk
from tkinter import Toplevel,Frame
from tkinter import ttk
from tkcalendar import Calendar
from PIL import Image, ImageTk
import sys
import platform
import threading
from datetime import datetime, timedelta
import subprocess
from tkinter import messagebox
import re


users = {"test-user1":"user-pass1",
         "test-user2":"user-pass2",
         "test-user3":"user-pass3",
         "1":"1",
}


def validate_credentials(username,password ):
    try:
        if users[username] == password:
            try:
                # Create a session with the credentials
                session = boto3.Session(
                    aws_access_key_id="some - id",
                    aws_secret_access_key="Some - key"
                )
                # Use S3 client to validate the credentials
                s3 = session.client('s3')
                s3.list_buckets()  # Try listing buckets to test the credentials
                return session  # Return the session if valid
            except Exception as e:
                print(e)
                return None
    except:
        print("Username/password not mapped")
    

def login_screen():
    def login():
        username = access_key_entry.get()
        password = secret_key_entry.get()
        # access_key = access_key_entry.get()
        # secret_key = secret_key_entry.get()
        
        if not username or not password:
            messagebox.showerror("Error", "Please enter both Username and Password")
            return
        
        global aws_session
        aws_session = validate_credentials(username, password)
        
        if aws_session:
            messagebox.showinfo("Success", "Login successful!")
            login_window.destroy()  # Close login screen
            main()
            
        else:
            messagebox.showerror("Error", "Invalid Credentials")

    # Create login window
    login_window = ctk.CTk()
    login_window.title("AWS Login")
    login_window.geometry("400x300")
    login_window.iconbitmap((resource_path(f"{script_dir}\\assets\\kowri-icon.ico")))

    # Access Key Input
    ctk.CTkLabel(login_window, text="Username:").pack(pady=10)
    access_key_entry = ctk.CTkEntry(login_window)
    access_key_entry.pack(pady=10)

    # Secret Key Input
    ctk.CTkLabel(login_window, text="Password:").pack(pady=10)
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

channel_path = ""


def run_query(channel_var: tk.StringVar, from_date: ctk.CTkEntry, to_date: ctk.CTkEntry, name_entry: ctk.CTkEntry,update_widgets,button):
    # s3_client = boto3.client("s3")
    s3_client = aws_session.client("s3")
    global channel_path

    unavailable_files = []
    bucket_name = 'all-kowri-datalake'
    current_directory = os.getcwd()
    channel_name = channel_var.get()
    if channel_name == "SecurePay-GH-Collections":
        channel_path = "KowriPartner"
    else:
        channel_path = "KowriBusiness"


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
        "InstantPayment-GH":"InstantPayment-GH",
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
        file_key = f'{channel_path}/{channel_name}/year={current_year}/month={current_month}/day={current_day}/{file_name_locs[channel_name]}_{current_year}_{current_month}_{current_day}.csv'
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
    
    if unavailable_files != []:
        user_response = messagebox.askokcancel(title="DataQuery", message="The following files could not be downloaded:\n" + "\n".join(unavailable_files) + "\nDo you wish to proceed?", icon="info")

        if not user_response:
                stop_event.set()
                update_widgets(button, "stopped")
            
        else:
            if not stop_event.is_set():
                try:
                    complete_file_df.to_csv(resource_path(f"{current_directory}\\data\\{output_file_name}.csv"), index=False)
                except Exception as e:
                    print(e)
                folder_path = resource_path(f"{current_directory}\\data")
                if platform.system() == "Windows":
                    os.startfile(folder_path)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", folder_path])
                else:
                    subprocess.Popen(["xdg-open", folder_path])
    else:
        if not stop_event.is_set():
                try:
                    complete_file_df.to_csv(resource_path(f"{current_directory}\\data\\{output_file_name}.csv"), index=False)
                except Exception as e:
                    print(e)
                folder_path = resource_path(f"{current_directory}\\data")
                if platform.system() == "Windows":
                    os.startfile(folder_path)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", folder_path])
                else:
                    subprocess.Popen(["xdg-open", folder_path])

# Global variable to store search results
search_results_df = pd.DataFrame()

def search_transaction(channel_var: tk.StringVar, from_date: ctk.CTkEntry, to_date: ctk.CTkEntry, 
                      transaction_ids: ctk.CTkEntry, status_label: ctk.CTkLabel, treeview: ttk.Treeview,
                      download_button: ctk.CTkButton):
    s3_client = boto3.client("s3")
    global channel_path
    global search_results_df
    
    unavailable_files = []
    bucket_name = 'all-kowri-datalake'
    channel_name = channel_var.get()
    transaction_ids_value = [id.strip() for id in transaction_ids.get().split(',')]

    channel_name = channel_var.get()
    if channel_name == "SecurePay-GH-Collections":
        channel_path = "KowriPartner"
    else:
        channel_path = "KowriBusiness"

    # Remove any empty strings
    transaction_ids_value = [id for id in transaction_ids_value if id]

    if not transaction_ids_value:
        status_label.configure(text="Please enter at least one transaction ID", text_color="red")
        # Disable download button since no results
        download_button.configure(state="disabled")
        return
    
    start_date = datetime.strptime(from_date.get(), "%d-%m-%Y")
    end_date = datetime.strptime(to_date.get(), "%d-%m-%Y")
    
    status_label.configure(text=f"Searching for {len(transaction_ids_value)} transaction(s)...", text_color="green")
    
    # Clear previous results
    for item in treeview.get_children():
        treeview.delete(item)
    
    # Clear existing columns
    for col in treeview["columns"]:
        treeview.heading(col, text='')
    
    # Reset the columns
    treeview["columns"] = ()
    treeview.column("#0", width=0, stretch=tk.NO)
    
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
        "InstantPayment-GH":"InstantPayment-GH",
    }
    
    
    found_transactions = []
    current_date = start_date
    
    while current_date <= end_date:
        if stop_event.is_set():
            print("Search stopped.")
            break
            
        print(f"Searching data for {current_date.strftime('%Y-%m-%d')}...")
        current_year = str(current_date.year)
        current_month = str(current_date.month).zfill(2)
        current_day = str(current_date.day).zfill(2)
        file_key = f'{channel_path}/{channel_name}/year={current_year}/month={current_month}/day={current_day}/{file_name_locs[channel_name]}_{current_year}_{current_month}_{current_day}.csv'
        file_name = f'{file_name_locs[channel_name]}_{current_year}_{current_month}_{current_day}.csv'
        print(file_key)
        print(file_name)
        
        try:
            response = s3_client.get_object(Bucket=bucket_name, Key=file_key)
            df = pd.read_csv(response['Body'])
            
            # Search for transaction IDs in all columns (case insensitive)
            # Create a mask to find rows containing any of the transaction IDs
            id_mask = df.astype(str).apply(
                lambda col: col.str.contains('|'.join(map(re.escape, transaction_ids_value)), case=False)
            ).any(axis=1)
            
            result_df = df[id_mask]
            
            if not result_df.empty:
                # Append the results to found transactions
                result_df['Source_File'] = file_name  # Add source file name for tracking
                found_transactions.append(result_df)
                print(f"Transaction(s) found in {file_name}")
                
        except Exception as e:
            unavailable_files.append(file_name)
            print(f"Couldn't search {file_name}: {str(e)}")
            
        current_date += timedelta(days=1)
    
    # Combine all found transactions
    if found_transactions:
        complete_result_df = pd.concat(found_transactions, ignore_index=True)
        
        # Store the results in the global variable for download
        search_results_df = complete_result_df
        
        # Get all columns from the dataframe
        all_columns = list(complete_result_df.columns)
        
        # Filter to only include columns that are in display_columns
        filtered_columns = [col for col in all_columns if col in display_columns]
        
        # If no columns match, fall back to all columns
        if not filtered_columns:
            filtered_columns = all_columns
        
        # Ensure 'Source_File' is included if not already in filtered columns
        if 'Source_File' not in filtered_columns:
            filtered_columns.append('Source_File')
        
        # Set the filtered columns to the treeview
        treeview["columns"] = filtered_columns
        
        # Configure column display
        treeview.column("#0", width=0, stretch=tk.NO)
        for col in filtered_columns:
            treeview.column(col, anchor=tk.W, width=100)
            treeview.heading(col, text=col, anchor=tk.W)
        
        # Add data to treeview
        for idx, row in complete_result_df.iterrows():
            values = [str(row[col]) for col in filtered_columns]
            treeview.insert("", tk.END, values=values)
        
        status_label.configure(
            text=f"Found {len(complete_result_df)} transaction(s) for {len(transaction_ids_value)} searched ID(s)", 
            text_color="green"
        )
        
        # Enable download button since results are available
        download_button.configure(state="normal")
    else:
        if not stop_event.is_set():
            status_label.configure(
                text=f"No transactions found for: {', '.join(transaction_ids_value)}", 
                text_color="red"
            )
            print("No transactions found in any files within the date range")
            # Disable download button since no results
            download_button.configure(state="disabled")
            # Reset the results dataframe
            search_results_df = pd.DataFrame()
    
    if unavailable_files and not stop_event.is_set():
        print(f"Couldn't search in these files: {unavailable_files}")

def download_search_results():
    global search_results_df
    
    if search_results_df.empty:
        messagebox.showinfo("DataQuery", "No search results to download.")
        return
    
    # Create data directory if it doesn't exist
    current_directory = os.getcwd()
    data_dir = resource_path(f"{current_directory}\\data")
    os.makedirs(data_dir, exist_ok=True)
    
    # Generate a filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"search_results_{timestamp}.csv"
    file_path = os.path.join(data_dir, filename)
    
    # Save to CSV
    try:
        search_results_df.to_csv(file_path, index=False)
        messagebox.showinfo("DataQuery", f"Results saved to {filename}")
        
        # Open the folder
        folder_path = resource_path(f"{current_directory}\\data")
        if platform.system() == "Windows":
            os.startfile(folder_path)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", folder_path])
        else:
            subprocess.Popen(["xdg-open", folder_path])
            
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save results: {str(e)}")


def main():
    def open_date_picker(date_entry: ctk.CTkEntry, root: ctk.CTk):
        date_window = Toplevel(root)
        date_window.title("Select Date")
        date_window.geometry("350x350")


        cal = Calendar(date_window, selectmode="day", year=datetime.now().year, month=datetime.now().month, day=datetime.now().day, date_pattern="dd-mm-yyyy")
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
    root.iconbitmap((resource_path(f"{script_dir}\\assets\\kowri-icon.ico")))
    root.option_add("*tearOff", False)
    root._set_appearance_mode("light")
    root.geometry("1000x650")  # Increased window size for better layout

    # Set up the style for ttk widgets
    style = ttk.Style(root)
    root.tk.call(
        "source",
        resource_path(f"{script_dir}\\assets\\Forest-ttk-theme-master\\forest-light.tcl"),
    )
    style.theme_use("forest-light")
    
    # Main container
    main_container = ctk.CTkFrame(root, fg_color="white")
    main_container.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Left sidebar for menu
    sidebar = ctk.CTkFrame(main_container, fg_color="#f0f0f0", width=200)
    sidebar.pack(side="left", fill="y", padx=(0, 10), pady=0)
    sidebar.pack_propagate(False)  # Prevent the frame from shrinking
    
    # Logo for sidebar
    logo_img_small = ctk.CTkImage(
        Image.open(resource_path(f"{script_dir}\\assets\\KowriLogo.png")),
        size=(150, 50),
    )
    sidebar_logo = ctk.CTkLabel(sidebar, image=logo_img_small, text="", fg_color="#f0f0f0")
    sidebar_logo.pack(pady=(20, 30))
    
    # Content area
    content_frame = ctk.CTkFrame(main_container, fg_color="white")
    content_frame.pack(side="right", fill="both", expand=True)
    
    # Menu frames
    download_frame = ctk.CTkFrame(content_frame, fg_color="white")
    search_frame = ctk.CTkFrame(content_frame, fg_color="white")
    
    # Function to switch between frames
    def show_frame(frame):
        download_frame.pack_forget()
        search_frame.pack_forget()
        frame.pack(fill="both", expand=True)
        
        # Update button styles
        if frame == download_frame:
            download_button.configure(fg_color="green", hover_color="#097969")
            search_button.configure(fg_color="#f0f0f0", hover_color="#e0e0e0", text_color="black")
        else:
            download_button.configure(fg_color="#f0f0f0", hover_color="#e0e0e0", text_color="black")
            search_button.configure(fg_color="green", hover_color="#097969")
    
    # Menu buttons
    download_button = ctk.CTkButton(
        sidebar, 
        text="Download Data", 
        command=lambda: show_frame(download_frame), 
        fg_color="green", 
        hover_color="#097969", 
        corner_radius=0,
        height=45,
        anchor="w"
    )
    download_button.pack(fill="x", pady=(0, 2))
    
    search_button = ctk.CTkButton(
        sidebar, 
        text="Search Data", 
        command=lambda: show_frame(search_frame), 
        fg_color="#f0f0f0", 
        hover_color="#e0e0e0", 
        text_color="black",
        corner_radius=0,
        height=45,
        anchor="w"
    )
    search_button.pack(fill="x")
    
    # Common channel options
    channels = [
        "MTN-GH-Collections", "MTN-GH-Disbursements", "Vodafone-GH-Collections", "Vodafone-GH-Disbursements", "Card-GH-GTMPGS", "InstantPayment-GH",
        "Card-GH-NGENIUS", "SecurePay-GH-Collections", "SecurePay-GH-Disbursements", "KBPlatform-MerchantOrder", "KBPlatform-Transaction"
    ]
    
    ### DOWNLOAD FRAME ###
    # Set up the Download frame
    logo_img = ctk.CTkImage(
        Image.open(resource_path(f"{script_dir}\\assets\\KowriLogo.png")),
        size=(200, 70),
    )

    download_title = ctk.CTkLabel(download_frame, text="Download Data", font=("Helvetica", 20, "bold"), text_color="green")
    download_title.pack(pady=(20, 30))

    channel_var_download = tk.StringVar(value=channels[0])

    channel_label = ctk.CTkLabel(download_frame, text="Select Channel:", text_color="green")
    channel_label.pack(pady=(0, 5))
    channel_dropdown = ctk.CTkOptionMenu(download_frame, variable=channel_var_download, values=channels, fg_color="white", button_color="green", dropdown_fg_color="white", text_color="black", dropdown_text_color="black", dropdown_hover_color="green")
    channel_dropdown.pack(pady=5)

    dates_frame = ctk.CTkFrame(download_frame, fg_color="white", bg_color="white")
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

    outName_frame = ctk.CTkFrame(download_frame, fg_color="white")
    outName_frame.pack(pady=(20, 5))
    outName_label = ctk.CTkLabel(outName_frame, text="Output File Name:", text_color="green")
    outName_label.pack(pady=(0, 5))
    outName_entry = ctk.CTkEntry(outName_frame, fg_color="white", bg_color="white", text_color="black", width=300)
    outName_entry.pack(pady=5)

    status_label = ctk.CTkLabel(download_frame, text="", text_color="green")
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

    def run_query_thread(channel_var, from_entry, to_entry, button, name_entry, update_widgets=update_widgets):
        stop_event.clear()
        update_widgets(button, "running")

        run_query(channel_var=channel_var, from_date=from_entry, to_date=to_entry, name_entry=name_entry, update_widgets=update_widgets, button=button)

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

    submit_button = ctk.CTkButton(download_frame, text="Run Query", command=lambda: check_data_validity(channel_var_download, from_entry, to_entry, submit_button, outName_entry), text_color="white", fg_color="green")
    submit_button.pack(pady=(30, 10))
    
    ### SEARCH FRAME ###
    # Set up the Search frame
    search_title = ctk.CTkLabel(search_frame, text="Search Data", font=("Helvetica", 20, "bold"), text_color="green")
    search_title.pack(pady=(20, 20))
    
    # Channel selection for search
    search_top_frame = ctk.CTkFrame(search_frame, fg_color="white")
    search_top_frame.pack(fill="x", padx=20, pady=10)
    
    channel_var_search = tk.StringVar(value=channels[0])
    search_channel_label = ctk.CTkLabel(search_top_frame, text="Select Channel:", text_color="green")
    search_channel_label.pack(side="left", padx=(0, 10))
    search_channel_dropdown = ctk.CTkOptionMenu(search_top_frame, variable=channel_var_search, values=channels, 
                                              fg_color="white", button_color="green", dropdown_fg_color="white", 
                                              text_color="black", dropdown_text_color="black", dropdown_hover_color="green",
                                              width=200)
    search_channel_dropdown.pack(side="left")
    
    # Date range selection
    search_dates_frame = ctk.CTkFrame(search_frame, fg_color="white")
    search_dates_frame.pack(fill="x", padx=20, pady=10)
    
    search_from_label = ctk.CTkLabel(search_dates_frame, text="From Date:", text_color="green")
    search_from_label.pack(side="left", padx=(0, 5))
    search_from_entry = ctk.CTkEntry(search_dates_frame, fg_color="white", bg_color="white", text_color="black", width=120)
    search_from_entry.pack(side="left", padx=(0, 5))
    search_from_button = ctk.CTkButton(search_dates_frame, text="Select", 
                                     command=lambda: open_date_picker(search_from_entry, root), 
                                     text_color="white", fg_color="green",
                                     width=70)
    search_from_button.pack(side="left", padx=(0, 20))
    
    search_to_label = ctk.CTkLabel(search_dates_frame, text="To Date:", text_color="green")
    search_to_label.pack(side="left", padx=(0, 5))
    search_to_entry = ctk.CTkEntry(search_dates_frame, fg_color="white", bg_color="white", text_color="black", width=120)
    search_to_entry.pack(side="left", padx=(0, 5))
    search_to_button = ctk.CTkButton(search_dates_frame, text="Select", 
                                   command=lambda: open_date_picker(search_to_entry, root), 
                                   text_color="white", fg_color="green",
                                   width=70)
    search_to_button.pack(side="left")
    
    # Transaction ID search
    search_id_frame = ctk.CTkFrame(search_frame, fg_color="white")
    search_id_frame.pack(fill="x", padx=20, pady=10)
    
    transaction_label = ctk.CTkLabel(search_id_frame, text="Transaction ID:", text_color="green")
    transaction_label.pack(side="left", padx=(0, 10))
    
    transaction_entry = ctk.CTkEntry(search_id_frame, fg_color="white", bg_color="white", text_color="black", width=300)
    transaction_entry.pack(side="left", padx=(0, 10), fill="x", expand=True)
    
    search_button_frame = ctk.CTkFrame(search_frame, fg_color="white")
    search_button_frame.pack(fill="x", padx=20, pady=10)
    
    search_status_label = ctk.CTkLabel(search_frame, text="", text_color="green")
    search_status_label.pack(pady=5)
    
    # Create a frame for the treeview
    tree_container = ctk.CTkFrame(search_frame, fg_color="white")
    tree_container.pack(fill="both", expand=True, padx=20, pady=10)
    
    # Create a frame for the treeview with scrollbars
    tree_frame = Frame(tree_container, bg="white")
    tree_frame.pack(fill="both", expand=True)
    
    # Create horizontal and vertical scrollbars
    tree_scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical")
    tree_scrollbar_y.pack(side="right", fill="y")
    
    tree_scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal")
    tree_scrollbar_x.pack(side="bottom", fill="x")
    
    # Create the treeview
    search_treeview = ttk.Treeview(tree_frame, yscrollcommand=tree_scrollbar_y.set, xscrollcommand=tree_scrollbar_x.set)
    search_treeview.pack(fill="both", expand=True)
    
    # Configure the scrollbars
    tree_scrollbar_y.config(command=search_treeview.yview)
    tree_scrollbar_x.config(command=search_treeview.xview)
    
    # Style the treeview for better appearance
    style.configure("Treeview", font=('Helvetica', 10), rowheight=25)
    style.configure("Treeview.Heading", font=('Helvetica', 10, 'bold'))
    
    # Create download button (initially disabled)
    download_results_button = ctk.CTkButton(
        search_button_frame, 
        text="Download Results", 
        command=download_search_results,
        text_color="white", 
        fg_color="green",
        state="disabled"
    )
    download_results_button.pack(side="right", padx=10)
    
    def search_thread(channel_var, from_entry, to_entry, transaction_entry, status_label, treeview, download_button):
        stop_event.clear()
        threading.Thread(target=search_transaction, args=(
            channel_var, from_entry, to_entry, transaction_entry, status_label, treeview, download_button
        )).start()
    
    # Search button
    search_btn = ctk.CTkButton(search_button_frame, text="Search Transaction", 
                                command=lambda: search_thread(
                                    channel_var_search, 
                                    search_from_entry, 
                                    search_to_entry, 
                                    transaction_entry,
                                    search_status_label,
                                    search_treeview,
                                    download_results_button
                                ), 
                                text_color="white", fg_color="green")
    search_btn.pack(side="left", pady=5)
    
    # Show the download frame by default
    show_frame(download_frame)
    
    # Start the main loop
    root.mainloop()

if __name__ == "__main__":
    aws_session = None
    login_screen()
    # main()
