import subprocess
import sys

def restart_program():
    python = sys.executable  # Get the path to the Python interpreter
    subprocess.Popen([python] + sys.argv)  # Restart the program with the same command line arguments
    sys.exit()  # Exit the current instance

# if __name__ == "__main__":
#     print("This is the original program.")
    
#     # Restart the program
#     restart_program()
