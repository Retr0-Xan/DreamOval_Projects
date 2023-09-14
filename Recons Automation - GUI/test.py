import tkinter as tk
from tkinter import ttk

def simulate_loading(loading_window):
    label = tk.Label(loading_window, text="Loading, please wait...")
    label.pack(pady=10)

    progress_bar = ttk.Progressbar(loading_window, mode="indeterminate")
    progress_bar.pack()

    # Function to update the progress bar
    def update_progress():
        if progress_bar['value'] < 100:
            progress_bar.step(10)  # Adjust the step as needed
            loading_window.after(100, update_progress)  # Update every 100 milliseconds

    update_progress()  # Start updating the progress bar

    # Simulate a time-consuming task (e.g., fetching data)
    import time
    time.sleep(5)  # Simulate 5 seconds of loading

    # Close the loading window when done
    loading_window.destroy()

def perform_action():
    # Simulate some action (e.g., data processing)
    import time
    time.sleep(3)  # Simulate 3 seconds of action

def start_loading_and_action():
    # Create the loading window
    loading_window = tk.Toplevel(root)
    loading_window.title("Loading...")

    # Run simulate_loading in a separate thread
    import threading
    loading_thread = threading.Thread(target=simulate_loading, args=(loading_window,))
    loading_thread.start()

    # Perform the action in the main thread
    perform_action()

root = tk.Tk()
root.title("Tkinter Loading Example")

button = tk.Button(root, text="Start Loading", command=start_loading_and_action)
button.pack(pady=20)

root.mainloop()
