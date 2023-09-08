import tkinter as tk
from tkinter import PhotoImage
from PIL import Image, ImageTk

def main():
    root = tk.Tk()
    root.title("Image in Label Example")

    # Load the image
    img = Image.open("C:\\Users\\Mark\\repos\DreamOval_Projects\Recons Automation - GUI\\assets\\login-bg(1).png")  

    # Get the dimensions of the label
    label_width = 500  # Replace with the desired label width
    label_height = 550  # Replace with the desired label height

    # Resize the image to fit the label
    img = img.resize((label_width, label_height))

    # Create a PhotoImage from the resized image
    img = ImageTk.PhotoImage(img)

    # Create a label with the image
    label = tk.Label(root, image=img)
    label.pack()

    # Keep a reference to the image to prevent it from being garbage collected
    label.img = img

    root.mainloop()

if __name__ == "__main__":
    main()
