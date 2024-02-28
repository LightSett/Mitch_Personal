import os
import zipfile
import tkinter as tk
from tkinter import filedialog

def unzip_all_zips(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.zip'):
                zip_path = os.path.join(root, file)
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(root)
                os.remove(zip_path)

def browse_directory():
    selected_directory = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(tk.END, selected_directory)

def run_extraction():
    directory = entry.get()
    if directory:
        unzip_all_zips(directory)
        status_label.config(text="Extraction completed.")
    else:
        status_label.config(text="Please select a directory.")

root = tk.Tk()
root.title("Zip Extractor")

label = tk.Label(root, text="Select a directory:")
label.pack()

entry = tk.Entry(root, width=50)
entry.pack()

browse_button = tk.Button(root, text="Browse", command=browse_directory)
browse_button.pack()

run_button = tk.Button(root, text="Run", command=run_extraction)
run_button.pack()

status_label = tk.Label(root, text="")
status_label.pack()

close_button = tk.Button(root, text="Close", command=root.destroy)
close_button.pack()

root.mainloop()
