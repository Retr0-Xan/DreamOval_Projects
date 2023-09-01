import customtkinter as ctk
from tkinter import ttk
root = ctk.CTk()
root.title("KowRecons")
root.option_add("*tearOff", False)
root._set_appearance_mode("light")

def create_recons_frame(tab):
    recons_frame = ctk.CTkFrame(tab)
    
    ctk.CTkLabel(recons_frame, text="Recons for yesterday?").pack()
    
    recons_var = ctk.StringVar()
    recons_var.set("yes")
    
    ttk.Radiobutton(recons_frame, text="Yes", variable=recons_var, value="yes").pack(anchor=ctk.W)
    ttk.Radiobutton(recons_frame, text="No", variable=recons_var, value="no").pack(anchor=ctk.W)
    
    recons_frame.pack(fill="both", expand=True)


tab_control = ttk.Notebook(root)

home_tab = ttk.Frame(tab_control)
recons_tab = ttk.Frame(tab_control)
logs_tab = ttk.Frame(tab_control)

tab_control.add(home_tab, text="Home")
tab_control.add(recons_tab, text="Recons")
tab_control.add(logs_tab, text="Logs")

create_recons_frame(recons_tab)

tab_control.pack(side=ctk.LEFT, fill="both", expand=True)

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