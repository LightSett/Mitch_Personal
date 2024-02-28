import os
import tkinter as tk
from tkinter import messagebox

# Define the root directory where you want to start the search
root_directory = "S:\Customers\23 V1Fiber\23.065.001 V1 Corning Lima\Scranton\LLD\Route 4"

# Function to search for the "LLD_Design.gdb" folder
def find_gdb_folder(start_directory):
    for root, dirs, files in os.walk(start_directory):
        if "LLD_Design.gdb" in dirs:
            return os.path.join(root, "LLD_Design.gdb")

# Search for the folder
found_folder = find_gdb_folder(root_directory)

if found_folder:
    message = f"The 'LLD_Design.gdb' folder was found at:\n{found_folder}"
    messagebox.showinfo("LLD_Design.gdb Found", message)
else:
    messagebox.showerror("LLD_Design.gdb Not Found", "The 'LLD_Design.gdb' folder was not found in the specified directory or its subdirectories.")
