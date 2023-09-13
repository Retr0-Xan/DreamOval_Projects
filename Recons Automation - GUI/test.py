import tkinter as tk
from tkinter import ttk
import sys

class Console(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.text_widget = tk.Text(self, wrap=tk.WORD)
        self.text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.scrollbar = tk.Scrollbar(self, command=self.text_widget.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.text_widget.config(yscrollcommand=self.scrollbar.set)
        
        sys.stdout = self
        
    def write(self, text):
        self.text_widget.insert(tk.END, text)
        self.text_widget.see(tk.END)  # Automatically scroll to the end
        
    def flush(self):
        pass

def main():
    root = tk.Tk()
    root.title("Console Redirect Example")

    console_frame = ttk.Frame(root)
    console_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    console = Console(console_frame)
    console.pack(fill=tk.BOTH, expand=True)

    # Test it by printing to the console
    print("Hello, this is console output.Lezzzgooooo")
    print("You can redirect stdout to this console.")

    root.mainloop()

if __name__ == "__main__":
    main()
