import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
import datetime
root = ctk.CTk()
root.title("KowRecons")
root.option_add("*tearOff", False)
root._set_appearance_mode("light")
root.geometry("500x550")


def render_recons_frame(tab):
    recons_frame = ctk.CTkFrame(tab,fg_color="white")
    


    yesterday_frame = ttk.LabelFrame(recons_frame,text="Set Recons Date",width=500,height=400)
    yesterday_frame.pack()
    yesterday_frame.pack_propagate(False)
    
    yesterady_msg = ctk.CTkLabel(yesterday_frame, text="Recons for yesterday?",text_color="green")
    yesterady_msg.pack()

    recons_var = ctk.StringVar()
    recons_var.set("yes")

    # def renderCalendar():


    def specifyDay():
        if recons_var.get() == "yes":
            print("you selectesd yes")
            dayCombo.current()
        elif recons_var.get() == "no":
            print()
  
    day_value_list = []
    for i in range(1,32):
        day_value_list.append(i)
    
    yesRadio = ttk.Radiobutton(yesterday_frame, text="Yes", variable=recons_var, value="yes",command=specifyDay)
    yesRadio.pack(anchor=ctk.W)
    noRadio = ttk.Radiobutton(yesterday_frame, text="No", variable=recons_var, value="no",command=specifyDay)
    noRadio.pack(anchor=ctk.W)


    dayLabel = ctk.CTkLabel(yesterday_frame,text="Day")
    dayLabel.pack()
    dayCombo = ttk.Combobox(yesterday_frame, state="disabled", values=day_value_list)
    dayCombo.current(0)
    dayCombo.pack()

    
    recons_frame.pack(fill="both", expand=True)

  

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


tab_control = ttk.Notebook(root)

home_tab = ttk.Frame(tab_control)
recons_tab = ttk.Frame(tab_control)
logs_tab = ttk.Frame(tab_control)

tab_control.add(home_tab, text="Home")
tab_control.add(recons_tab, text="Recons")
tab_control.add(logs_tab, text="Logs")

render_recons_frame(recons_tab)

tab_control.pack(side=ctk.LEFT, fill="both", expand=True)

root.mainloop()