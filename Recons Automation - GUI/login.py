import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
root = ctk.CTk()
root.title("KowRecons")
root.option_add("*tearOff", False)
root._set_appearance_mode("light")
root.geometry("500x550")

welcome_frame = ctk.CTkFrame(root,fg_color="white")
# welcome_frame.pack_propagate(False)
welcome_frame.pack()

welcome_msg = ctk.CTkLabel(welcome_frame,text="Welcome Back!",fg_color="#ebf0ec",text_color="Green",font=ctk.CTkFont("Rounded Arial",size = 40,weight="bold"))
# welcome_msg.pack_propagate(False)
welcome_msg.pack()

signin_msg = ctk.CTkLabel(welcome_frame,text="Select User",fg_color="#ebf0ec",text_color="Green",font=ctk.CTkFont("Rounded Arial",size = 30,weight="bold"))
# welcome_msg.pack_propagate(False)
signin_msg.pack()





root.mainloop()