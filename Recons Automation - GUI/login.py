import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk


def main():
    root = ctk.CTk()
    root.title("KowRecons")
    root.option_add("*tearOff", False)
    root._set_appearance_mode("light")
    root.geometry("900x550")

    bg_img=Image.open("C:\\Users\\Mark\\repos\DreamOval_Projects\Recons Automation - GUI\\assets\\login-bg(1).png")
    label_width = root.winfo_screenwidth()
    label_height = root.winfo_screenheight()
    bg_img = bg_img.resize((label_width, label_height))
    bg_img = ImageTk.PhotoImage(bg_img)

    bg_label = ctk.CTkLabel(root, image=bg_img)
    bg_label.pack(fill="both",expand=True)

    bg_label.bg_img = bg_img


    # welcome_frame = ctk.CTkFrame(bg_label,fg_color="white")
    # # welcome_frame.pack_propagate(False)
    # welcome_frame.pack()

    # welcome_msg = ctk.CTkLabel(welcome_frame,text="Welcome Back!",fg_color="#ebf0ec",text_color="Green",font=ctk.CTkFont("Rounded Arial",size = 40,weight="bold"))
    # # welcome_msg.pack_propagate(False)
    # welcome_msg.pack()

    # signin_msg = ctk.CTkLabel(welcome_frame,text="Select User",fg_color="#ebf0ec",text_color="Green",font=ctk.CTkFont("Rounded Arial",size = 30,weight="bold"))
    # # welcome_msg.pack_propagate(False)
    # signin_msg.pack()





    root.mainloop()

if __name__ == "__main__":
    main()