import tkinter as tk
from tkinter import ttk

def set_combobox_value():
    combo.set("2")  # Set the value of the Combobox to the integer 2

root = tk.Tk()
root.title("Set Combobox Value Example")

combo = ttk.Combobox(root, values=[1, 2, 3])
combo.pack(padx=10, pady=10)

button = tk.Button(root, text="Set Value", command=set_combobox_value)
button.pack()

root.mainloop()
