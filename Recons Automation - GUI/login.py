import customtkinter as ctk
import tkinter as tk
from tkinter import ttk,PhotoImage
from PIL import Image, ImageTk


def main():

    def on_map(event):
            root.state("zoomed")


    root = ctk.CTk()
    root.geometry(f"{root.winfo_screenwidth()} x {root.winfo_screenheight()}")
    root.bind("<Map>", on_map)
    root.option_add("*tearOff", False)
    root._set_appearance_mode("light")
    root.title("KowRecons")
    bg_img = ctk.CTkImage(Image.open("C:\\Users\\Mark\\repos\DreamOval_Projects\Recons Automation - GUI\\assets\\login-bg(4).png"), size = (root.winfo_screenwidth(),root.winfo_screenheight()))
    bg_label = ctk.CTkLabel(root, image=bg_img,text="")
    bg_label.pack(fill="both",expand=True)

    bg_label.bg_img = bg_img

    logo_img = ctk.CTkImage(Image.open("C:\\Users\\Mark\\repos\DreamOval_Projects\Recons Automation - GUI\\assets\\KowriLogo.png"),size=(200,70))
    label = ctk.CTkLabel(bg_label, image=logo_img,text="",fg_color="#f0f2f1")
    label.place(relx=0.02,rely=0.06)


    # welcome_frame = ctk.CTkFrame(bg_label,fg_color="#01911e",corner_radius=10,bg_color="#ebf0ec",width=400,height=300)
    # welcome_frame.pack_propagate(False)
    # welcome_frame.place(relx=0.36,rely=0.05)

    welcome_msg = ctk.CTkLabel(bg_label,text="Welcome Back!",text_color="White",bg_color= "black",font=ctk.CTkFont("Rounded Arial",size = 40,weight="normal"))
    welcome_msg.pack_propagate(False)
    welcome_msg.place(relx=0.36,rely=0.05)


    # user_frame = ctk.CTkFrame(bg_label,width=400,height=400,fg_color="green")
    # user_frame.pack()





    root.mainloop()

if __name__ == "__main__":
    main()