import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
import calendar
from datetime import datetime,date,timedelta
from PIL import Image, ImageTk


def main ():
    root = ctk.CTk()
    root.title("KowRecons")
    root.option_add("*tearOff", False)
    root._set_appearance_mode("light")
    root.geometry("500x550")


    def render_recons_frame(tab):
        recons_frame = ctk.CTkFrame(tab,fg_color="white")

        logo_img = ctk.CTkImage(Image.open("C:\\Users\\Mark\\repos\DreamOval_Projects\Recons Automation - GUI\\assets\\KowriLogo.png"),size=(200,70))
        label = ctk.CTkLabel(recons_frame, image=logo_img,text="",fg_color="white")
        label.pack()
        


        yesterday_frame = ttk.LabelFrame(recons_frame,text="Set Recons Date",width=500,height=400)
        yesterday_frame.pack()
        yesterday_frame.pack_propagate(False)
        
        yesterady_msg = ctk.CTkLabel(yesterday_frame, text="Recons for yesterday?",text_color="green")
        yesterady_msg.pack()

        recons_var = ctk.StringVar()
        recons_var.set("yes")



        def yesterdayRecons():
            dayCombo.configure(state="disabled")
            monthCombo.configure(state="disabled")
            yearCombo.configure(state="disabled")
            today = datetime.today().strftime('%d_%b_%y')

            # Yesterday's date
            yesterday = (datetime.now() - timedelta(1)).strftime('_%d %b_%y')
            print(yesterday)

            # Current month
            current_month = calendar.month_abbr[date.today().month].upper()
            print(current_month)

            recons_yesterday = str((datetime.now() - timedelta(1)).strftime('%Y-%m-%d') + ' 00:00:00')
            print(recons_yesterday)

            if date.today().month < 10:
                GIPmonth = "0" + str(date.today().month)
            else:
                GIPmonth = str(date.today().month)

            if date.today().day < 10:
                GIPday = "0" + str(date.today().day + 1)
            else:
                GIPday = str(date.today().day + 1)

            GIPdate = str(str(date.today().year) + str(GIPmonth) + str((GIPday)))
            print(GIPdate)

        def customDateRecons():
                dayCombo.configure(state="enabled")
                monthCombo.configure(state="enabled")
                yearCombo.configure(state="enabled")


        def displayValues():
            if recons_var.get() == "no":
                customDay = dayCombo.get()
                customMonth = monthCombo.get()
                customYear = yearCombo.get()[-2:]
                GIPday = int(customDay) + 1
                if int(customDay) < 10:
                    customDay = "0" + customDay
                month_name = calendar.month_abbr[int(customMonth)]
                if int(customMonth) < 10:
                    customMonth = "0" + customMonth
                yesterday = "_" + customDay + " " + month_name + "_" + customYear
                print(yesterday)

                recons_yesterday = "20" + customYear + "-" + customMonth + "-" + customDay + ' 00:00:00'
                print(recons_yesterday)

                if int(customMonth) < 10:
                    GIPmonth = customMonth

                if int(GIPday) < 10:
                    GIPday = "0" + str(GIPday)
                GIPdate = str("20" + str(customYear) + str(GIPmonth) + str((GIPday)))
                print(GIPdate)

            elif recons_var.get() == "yes":
                yesterdayRecons()
    
        day_value_list = []
        for i in range(1,32):
            day_value_list.append(i)

        month_value_list = []
        for i in range(1,13):
            month_value_list.append(i)

        
        yesRadio = ttk.Radiobutton(yesterday_frame, text="Yes", variable=recons_var, value="yes",command=yesterdayRecons)
        yesRadio.pack(anchor=ctk.W)
        noRadio = ttk.Radiobutton(yesterday_frame, text="No", variable=recons_var, value="no",command=customDateRecons)
        noRadio.pack(anchor=ctk.W)


        dayLabel = ctk.CTkLabel(yesterday_frame,text="Day")
        dayLabel.pack()
        dayCombo = ttk.Combobox(yesterday_frame, state="disabled", values=day_value_list)
        dayCombo.current(0)
        dayCombo.pack()

        monthLabel = ctk.CTkLabel(yesterday_frame,text="Month")
        monthLabel.pack()
        monthCombo = ttk.Combobox(yesterday_frame, state="disabled", values=month_value_list)
        monthCombo.current(0)
        monthCombo.pack()

        yearLabel = ctk.CTkLabel(yesterday_frame,text="Year")
        yearLabel.pack()
        yearCombo = ttk.Combobox(yesterday_frame, state="disabled", values=["2022","2023"])
        yearCombo.current(0)
        yearCombo.pack()

        startButton = ctk.CTkButton(yesterday_frame,text="Start",fg_color="green",hover_color="#1bcf48",command=lambda:(displayValues()))
        startButton.pack(pady=(20,0))


        
        recons_frame.pack(fill="both", expand=True)

    
    # Make the window interactive
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

if __name__ == "__main__":
    main()