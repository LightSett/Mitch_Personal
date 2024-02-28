import tkinter as tk
from tkinter import filedialog, messagebox
import arcpy
import os
import pandas as pd
import datetime
import shutil

def close_window():
    root.destroy()

def run_export():
    if not gdb_list:
        messagebox.showwarning("No GDBs", "Please load GDBs or GDB List first.")
        return
    
    create_output_folder()
    
    print("Exporting MakeReadySummary attribute tables...")
    for gdb in gdb_list:
        arcpy.env.workspace = gdb
        feature_classes = arcpy.ListFeatureClasses()
        
        if 'MakeReadySummary' in feature_classes:
            feature_path = os.path.join(gdb, 'MakeReadySummary')
            output_excel_file = os.path.join(output_folder, f"{os.path.basename(gdb)}_MakeReadySummary.xlsx")
            export_to_excel(feature_path, output_excel_file)
            message = f"Exported attribute table of MakeReadySummary from {os.path.basename(gdb)} to {output_excel_file}"
            messagebox.showinfo("Export Complete", message)
            print(message)
        else:
            print(f"No MakeReadySummary found in {os.path.basename(gdb)}")

    zip_output_folder()

def create_output_folder():
    global output_folder
    downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
    today_date_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_folder = os.path.join(downloads_folder, f"MRSummary_{today_date_time}")
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

def export_to_excel(input_table, output_excel):
    arcpy.TableToExcel_conversion(input_table, output_excel)

def zip_output_folder():
    shutil.make_archive(output_folder, 'zip', output_folder)
    messagebox.showinfo("Zip Complete", f"Exported files zipped to {output_folder}.zip")
    print(f"Exported files zipped to {output_folder}.zip")

def load_gdb():
    gdb_path = filedialog.askdirectory()
    if gdb_path:
        gdb_list.append(gdb_path)
        update_gdb_messagebox()

def load_gdb_list():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            if 'GDB_Path' in df.columns:
                for path in df['GDB_Path']:
                    gdb_list.append(path)
                update_gdb_messagebox()
                print("Loaded GDB paths from the Excel file.")
            else:
                messagebox.showerror("Invalid File", "Selected file does not contain 'GDB_Path' column.")
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {e}")

def update_gdb_messagebox():
    gdb_messagebox.delete(1.0, tk.END)
    gdb_messagebox.insert(tk.END, '\n'.join(gdb_list))

gdb_list = []
output_folder = ""

root = tk.Tk()
root.title("GDB Export Tool")

frame = tk.Frame(root)
frame.pack(padx=20, pady=20)

close_button = tk.Button(frame, text="Close", command=close_window)
close_button.pack(side=tk.RIGHT, padx=5)

run_button = tk.Button(frame, text="Run", command=run_export)
run_button.pack(side=tk.RIGHT, padx=5)

load_gdb_button = tk.Button(frame, text="Load GDB", command=load_gdb)
load_gdb_button.pack(side=tk.LEFT, padx=5)

load_gdb_list_button = tk.Button(frame, text="Load GDB List", command=load_gdb_list)
load_gdb_list_button.pack(side=tk.LEFT, padx=5)

gdb_messagebox = tk.Text(root, height=10, width=50)
gdb_messagebox.pack(padx=20, pady=10)

root.mainloop()
