import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk


def main():
    root = ctk.CTk()
    root.title("KowRecons")
    root.option_add("*tearOff", False)
    root._set_appearance_mode("light")
    root.geometry("1100x550")

    bg_img=Image.open("C:\\Users\\Mark\\repos\DreamOval_Projects\Recons Automation - GUI\\assets\\login-bg(1).png")
    label_width = root.winfo_screenwidth()
    label_height = root.winfo_screenheight()
    bg_img = bg_img.resize((label_width, label_height))
    bg_img = ImageTk.PhotoImage(bg_img)

    bg_label = ctk.CTkLabel(root, image=bg_img)
    bg_label.pack(fill="both",expand=True)

    bg_label.bg_img = bg_img

    # Load the PNG image using PIL
    image = Image.open("C:\\Users\\Mark\\repos\DreamOval_Projects\Recons Automation - GUI\\assets\\KowriLogo.png")

    # Convert the PIL Image to a PhotoImage
    img = ImageTk.PhotoImage(image.resize((200,70)))

    # Create a Canvas widget
    canvas = tk.Canvas(bg_label, width=200, height=70)
    canvas.place(relx=0.05,rely=0.06)

    # Create a Label on the Canvas to display the transparent image
    label = tk.Label(canvas, image=img)
    label.photo = img  # Keep a reference to the image to prevent it from being garbage collected

    # Place the Label on the Canvas
    canvas.create_window(0, 0, anchor=tk.NW, window=label)

    welcome_frame = ctk.CTkFrame(bg_label,fg_color="#01911e",corner_radius=10,bg_color="#ebf0ec",width=400,height=300)
    welcome_frame.pack_propagate(False)
    welcome_frame.place(relx=0.37,rely=0.1)

    welcome_msg = ctk.CTkLabel(welcome_frame,text="Welcome Back!",text_color="White",font=ctk.CTkFont("Rounded Arial",size = 40,weight="bold"))
    welcome_msg.pack_propagate(False)
    welcome_msg.pack()





    root.mainloop()

if __name__ == "__main__":
    main()