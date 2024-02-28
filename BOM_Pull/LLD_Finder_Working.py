import os
import tkinter as tk
from tkinter import messagebox, filedialog
from openpyxl import Workbook
from datetime import datetime

def is_archive_folder(folder_name):
    return "archive" in folder_name.lower()

def find_aprx_modification_date(folder_path):
    for file in os.listdir(folder_path):
        if file.endswith(".aprx"):
            aprx_file_path = os.path.join(folder_path, file)
            modification_date = os.path.getmtime(aprx_file_path)
            return modification_date
    return None

def find_gdb_folders(start_directory):
    found_folders = []
    target_folder = "LLD_Design.gdb"
    for root, dirs, files in os.walk(start_directory):
        if is_archive_folder(root):
            continue

        lowercase_dirs = [dir.lower() for dir in dirs]  # Convert directory names to lowercase
        if target_folder.lower() in lowercase_dirs:
            found_folder = os.path.join(root, target_folder)
            aprx_modification_date = find_aprx_modification_date(root)
            if aprx_modification_date:
                found_folders.append((aprx_modification_date, found_folder))
    return found_folders

def browse_directory():
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_path.delete(0, tk.END)
        entry_path.insert(tk.END, folder_path)

def run_script():
    directory_path = entry_path.get()
    if directory_path:
        found_folders = find_gdb_folders(directory_path)

        if found_folders:
            wb = Workbook()
            sheet = wb.active
            sheet.title = "LLD_List"
            sheet['A1'] = "Date"
            sheet['B1'] = "Found LLD_Design.gdb Folders:"
            
            for idx, (aprx_modification_date, folder_path) in enumerate(found_folders, start=2):
                date_formatted = datetime.fromtimestamp(aprx_modification_date).strftime("%Y-%m-%d %H:%M:%S")
                sheet[f'A{idx}'] = date_formatted
                sheet[f'B{idx}'] = folder_path

            folder_name = os.path.basename(directory_path)
            file_name = f"{folder_name}_LLD_List.xlsx"
            script_directory = os.path.dirname(os.path.abspath(__file__))
            excel_file_path = os.path.join(script_directory, file_name)
            wb.save(excel_file_path)

            messagebox.showinfo("Excel File Created", f"Excel file '{file_name}' created in the script directory.")
        else:
            messagebox.showerror("LLD_Design.gdb Not Found", "No 'LLD_Design.gdb' folders were found in the specified directory or its subdirectories (excluding folders containing 'archive' in their names).")
    else:
        messagebox.showerror("Directory Path Missing", "Please select a directory first.")

def close_window():
    root.destroy()

root = tk.Tk()
root.title("LLD Finder")

# Setting the background color to fluorescent pink
root.configure(bg="#ff00ff")

frame = tk.Frame(root, bg="#ff00ff")  # Set frame background color to match root
frame.pack(padx=10, pady=10)

label_instruction = tk.Label(frame, text="Paste or browse to the filepath \nof the parent folder where GDBs \nare contained.", font=("Goudy Stout", 12, "bold"), bg="#ff00ff", fg="#00ff00", highlightbackground="white")
label_instruction.pack(padx=10, pady=10)

input_frame = tk.Frame(frame, bg="#ff00ff")
input_frame.pack(padx=10, pady=10)

label_path = tk.Label(input_frame, text="Directory Path:", font=("Arial", 12, "bold"), bg="#ff00ff")
label_path.grid(row=0, column=0, sticky="w")

entry_path = tk.Entry(input_frame, width=40)
entry_path.grid(row=0, column=1, padx=5, pady=5)

button_browse = tk.Button(input_frame, text="Browse", font=("Arial", 12, "bold"), command=browse_directory)
button_browse.grid(row=0, column=2, padx=5, pady=5)

button_run = tk.Button(frame, text="Find Me some GDBs", font=("Arial", 12, "bold"), command=run_script)
button_run.pack(pady=10)

button_close = tk.Button(frame, text="Close", font=("Arial", 12, "bold"), command=close_window)
button_close.pack(pady=10)

root.mainloop()
