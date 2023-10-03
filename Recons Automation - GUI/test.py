import customtkinter as ctk
import tkinter as tk
from tkinter import ttk

root = ctk.CTk()

root.title("Test")
root.option_add("*tearOff", False)
root._set_appearance_mode("light")
root.geometry("500x550")

yesterday_frame = ttk.LabelFrame(root,text="Set Recons Date",width=500,height=400)
yesterday_frame.pack()
yesterday_frame.pack_propagate(False)

yesterady_msg = ctk.CTkLabel(yesterday_frame, text="Recons for yesterday?",text_color="green")
yesterady_msg.pack()

root.mainloop()