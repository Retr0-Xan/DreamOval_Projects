import customtkinter as ctk
from tkinter import ttk
root = ctk.CTk()
root.title("KowRecons")
root.option_add("*tearOff", False)
root._set_appearance_mode("light")

root.columnconfigure(index=0, weight=1)
root.columnconfigure(index=1, weight=1)
root.columnconfigure(index=2, weight=1)
root.rowconfigure(index=0, weight=1)
root.rowconfigure(index=1, weight=1)
root.rowconfigure(index=2, weight=1)

style = ttk.Style(root)
# Import the tcl file
root.tk.call("source", "Forest-ttk-theme-master/forest-light.tcl")

# Set the theme with the theme_use method
style.theme_use("forest-light")


root.mainloop()