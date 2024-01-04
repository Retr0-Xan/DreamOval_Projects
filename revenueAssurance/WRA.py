import pandas as pd
import customtkinter as ctk
import tkinter as tk
from tkinter import *
from tkinter import ttk
import csv
from tkinter import filedialog
from PIL import Image
import os, sys


def set_working_directory_to_script_location():
    if getattr(sys, "frozen", False):
        # We are running from a bundled executable
        script_dir = os.path.dirname(sys.executable)
    else:
        # We are running the script directly
        script_dir = os.path.dirname(os.path.abspath(__file__))

    os.chdir(script_dir)
    return script_dir


# Call the function to set the working directory
script_dir = set_working_directory_to_script_location()


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def choose_file():
    root.filename = filedialog.askopenfilename(
        initialdir="/.",
        title="open file",
        filetypes=(
            ("All Files", "*.*"),
            ("CSV Files", "*.csv"),
            ("Excel Files", "*.xlsx"),
        ),
    )
    return root.filename


def get_data():
    global all_items
    file = pd.read_excel(choose_file())
    merchants = file["merchantName"].drop_duplicates()
    merchants = merchants.tolist()

    highestComission = 0
    highestMerchant = ""
    for merchant in merchants:
        singleMerchantDf = file.loc[file["merchantName"] == merchant]

        count = len(singleMerchantDf)
        sum = singleMerchantDf["amount"].sum()

        cost = find_cost(df=singleMerchantDf)
        commission = find_commission(df=singleMerchantDf)

        if commission > highestComission:
            highestComission = commission
            highestMerchant = merchant

        merchant_data = (merchant, count, sum, cost, commission)
        table.insert(parent="", index=tk.END, values=merchant_data)

    all_items = [table.item(item) for item in table.get_children()]


def export_data(tree, filename):
    # Get column names
    columns = tree["columns"]

    # Open a CSV file for writing
    with open(filename, "w", newline="") as csvfile:
        # Create a CSV writer object
        csv_writer = csv.writer(csvfile)

        # Write header row with column names
        csv_writer.writerow(columns)

        # Write data rows
        for item in tree.get_children():
            # Get data for each row
            values = tuple(map(str, tree.item(item, "values")))

            # Write the row to the CSV file
            csv_writer.writerow(values)


def filter_data():
    filter_text = filter_var.get().lower()
    table.delete(*table.get_children())
    filtered_items = [
        item for item in all_items if filter_text in str(item["values"]).lower()
    ]
    for item in filtered_items:
        table.insert("", "end", values=item["values"])


def find_cost(df: pd.DataFrame):
    cost = 0.00
    for i in range(0, len(df)):
        if df.iloc[i]["paymentMode"] == "CASH":
            cost += 0.00
        elif df.iloc[i]["paymentMode"] == "VODAFONE_CASH":
            cost += 1 / 100 * (df.iloc[i]["amount"])
        elif (
            df.iloc[i]["paymentMode"] == "MASTERCARD"
            or df.iloc[i]["paymentMode"] == "VISA"
        ):
            cost += 2 / 100 * (df.iloc[i]["amount"])
        elif df.iloc[i]["paymentMode"] == "MTN_MONEY":
            if df.iloc[i]["amount"] < 1000:
                cost += 1 / 100 * (df.iloc[i]["amount"])
            elif df.iloc[i]["amount"] > 999:
                cost += 10.00
    return cost


def find_commission(df: pd.DataFrame):
    commission = 0.00
    for i in range(0, len(df)):
        commission += df.iloc[i]["markupFee"] + df.iloc[i]["commissionAmount"]

    return commission


root = ctk.CTk(fg_color="#e3e8e4")
root.title("Revenue Assurance Generator")
root.option_add("*tearOff", False)
root._set_appearance_mode("light")
root.geometry("1260x600")
style = ttk.Style(root)
root.tk.call(
    "source",
    resource_path(f"{script_dir}/assets/Forest-ttk-theme-master/forest-light.tcl"),
)
style.theme_use("forest-light")

merchants = []

menu_frame = ctk.CTkFrame(root, width=200, height=root.winfo_height(), fg_color="white")
menu_frame.pack_propagate(False)
menu_frame.pack(side="left", fill="y")


logo_img = ctk.CTkImage(
    Image.open(resource_path(f"{script_dir}/assets/KowriLogo.png")), size=(180, 60)
)
label = ctk.CTkLabel(menu_frame, image=logo_img, text="")
label.pack()

dashboard_img = ctk.CTkImage(
    Image.open(resource_path(f"{script_dir}/assets/monitor.png")), size=(38, 38)
)
history_img = ctk.CTkImage(
    Image.open(resource_path(f"{script_dir}/assets/history.png")), size=(38, 38)
)
export_img = ctk.CTkImage(
    Image.open(resource_path(f"{script_dir}/assets/export.png")), size=(15, 15)
)
plus_img = ctk.CTkImage(Image.open(resource_path(f"{script_dir}/assets/plus.png")), size=(15, 15))


dashboard_btn = ctk.CTkButton(
    menu_frame,
    text="Dashboard",
    text_color="black",
    hover_color="#7bed96",
    fg_color="#48db6b",
    height=70,
    corner_radius=0,
    font=ctk.CTkFont("Helvetica", size=20),
    image=dashboard_img,
    compound="left",
)
dashboard_btn.pack(pady=18, fill="x")
generate_btn = ctk.CTkButton(
    menu_frame,
    text="Generate",
    text_color="black",
    hover_color="#7bed96",
    fg_color="#48db6b",
    height=70,
    corner_radius=0,
    font=ctk.CTkFont("Helvetica", size=20),
)
generate_btn.pack(pady=18, fill="x")
history_btn = ctk.CTkButton(
    menu_frame,
    text="History",
    text_color="black",
    hover_color="#7bed96",
    fg_color="#48db6b",
    height=70,
    corner_radius=0,
    font=ctk.CTkFont("Helvetica", size=20),
    image=history_img,
    compound="left",
)
history_btn.pack(pady=18, fill="x")

option_frame = ctk.CTkFrame(root, fg_color="#e3e8e4", height=100, corner_radius=0)
option_frame.pack(fill="x", pady=15)
filter_var = tk.StringVar()
search_bar = ctk.CTkEntry(
    option_frame,
    width=350,
    height=10,
    fg_color="#d4d4d4",
    corner_radius=0,
    text_color="green",
    textvariable=filter_var,
    border_width=0,
)
search_bar.grid(column=0, row=0, padx=(250, 0))

search_btn = ctk.CTkButton(
    option_frame, text="Search", fg_color="green", width=15, command=filter_data
)
search_btn.grid(column=1, row=0, padx=(6, 30))

select_btn = ctk.CTkButton(
    option_frame,
    text="Select file",
    command=get_data,
    fg_color="#046924",
    text_color="white",
    bg_color="#046924",
    hover_color="#18ad48",
    image=plus_img,
    compound="left",
)
select_btn.grid(column=5, row=0, padx=(0, 30))

export_btn = ctk.CTkButton(
    option_frame,
    text="Export Data",
    fg_color="#046924",
    text_color="white",
    bg_color="#046924",
    hover_color="#18ad48",
    command=lambda: export_data(table, "RAG_(month).csv"),
    image=export_img,
    compound="left",
)
export_btn.grid(column=6, row=0)

table_frame = ctk.CTkFrame(root)
table_frame.pack()
table = ttk.Treeview(
    table_frame,
    columns=(
        "Merchants",
        "Total Volume",
        "Total Value",
        "Total Cost",
        "Total Comission",
    ),
    show="headings",
    height=20,
)
for col in table["columns"]:
    table.heading(col, text=col, anchor="center")
    table.column(col, anchor="center")
table.grid(column=0, row=5, columnspan=3)


root.mainloop()
