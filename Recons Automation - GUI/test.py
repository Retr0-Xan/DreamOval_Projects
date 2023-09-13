import tkinter as tk
from tkinter import ttk

def start_spinning():
    # Start the spinning wheel animation
    progress_bar.start()

def stop_spinning():
    # Stop the spinning wheel animation
    progress_bar.stop()

root = tk.Tk()
root.title("Spinning Wheel Example")

# Create a Progressbar with the "indeterminate" style
progress_bar = ttk.Progressbar(root, mode="indeterminate")
progress_bar.pack()

# Create buttons to start and stop the spinning animation
start_button = tk.Button(root, text="Start Spinning", command=start_spinning)
stop_button = tk.Button(root, text="Stop Spinning", command=stop_spinning)
start_button.pack()
stop_button.pack()

root.mainloop()
