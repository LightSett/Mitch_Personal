import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os

# Get the directory of the script hosting the GUI
script_directory = os.path.dirname(os.path.abspath(__file__))

def add_field():
    entry_text = entry.get()
    
    if os.path.exists(entry_text):
        if os.path.isfile(entry_text):
            field_list.insert(tk.END, f"File: {entry_text}")
        elif os.path.isdir(entry_text):
            field_list.insert(tk.END, f"Folder: {entry_text}")
        entry.delete(0, tk.END)
    else:
        messagebox.showerror("Invalid Path", "The specified path does not exist.")

def clear_fields():
    field_list.delete(0, tk.END)

def browse_file():
    file_path = filedialog.askopenfilename(initialdir=script_directory)
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

def browse_directory():
    dir_path = filedialog.askdirectory(initialdir=script_directory)
    if dir_path:
        entry.delete(0, tk.END)
        entry.insert(0, dir_path)
        add_field()  # Automatically add the selected directory path

def find_gdb_folder(start_directory):
    found_folders = []

    for root, dirs, files in os.walk(start_directory):
        # Check if any directory is named "archive" (case-insensitive)
        if any(d.lower() == "archive" for d in dirs):
            continue

        if "LLD_Design.gdb" in dirs:
            found_folders.append(os.path.join(root, "LLD_Design.gdb"))

    return found_folders

def search_and_display_gdb_folders():
    selected_items = field_list.get(0, tk.END)
    found_folders = []

    for item in selected_items:
        if item.startswith("Folder: "):
            folder_path = item[8:]
            found_folders.extend(find_gdb_folder(folder_path))

    if found_folders:
        message = "The 'LLD_Design.gdb' folders were found at:\n\n"
        message += "\n".join(found_folders)
        messagebox.showinfo("LLD_Design.gdb Search Result", message)
    else:
        messagebox.showinfo("LLD_Design.gdb Search Result", "The 'LLD_Design.gdb' folder was not found in any of the specified directories or their subdirectories.")

root = tk.Tk()
root.title("Field List GUI")

# Entry field
entry_label = tk.Label(root, text="Field Name or File/Folder Path:")
entry_label.pack()
entry = tk.Entry(root)
entry.pack()

# Add Field button
add_button = tk.Button(root, text="Add Field", command=add_field)
add_button.pack()

# Clear Fields button
clear_button = tk.Button(root, text="Clear Fields", command=clear_fields)
clear_button.pack()

# Browse File button
browse_file_button = tk.Button(root, text="Browse File", command=browse_file)
browse_file_button.pack()

# Browse Directory button
browse_dir_button = tk.Button(root, text="Browse Directory", command=browse_directory)
browse_dir_button.pack()

# Find LLD_Design.gdb button
find_gdb_button = tk.Button(root, text="Find LLD_Design.gdb", command=search_and_display_gdb_folders)
find_gdb_button.pack()

# List of Imported Fields
field_list_label = tk.Label(root, text="Imported Fields:")
field_list_label.pack()
field_list = tk.Listbox(root)
field_list.pack()

root.mainloop()
