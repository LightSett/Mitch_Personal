import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font

# Function to find and catalog "Archive" folders and their contents recursively
def find_archive_folders(root_dir, output_worksheet):
    archive_folders = []
    for foldername, subfolders, filenames in os.walk(root_dir):
        for subfolder in subfolders:
            if "Archive" in subfolder:
                folder_path = os.path.join(foldername, subfolder)

                # Check if the folder is a Geodatabase folder (ends with .gdb)
                if subfolder.lower().endswith(".gdb"):
                    # Exclude GDB folders and their contents
                    print(f"Excluded GDB folder: {folder_path}")
                    continue

                archive_folders.append((foldername, subfolder, folder_path))

                # Verbose terminal messaging
                print(f"Scanning directory: {foldername}")
                print(f"Archive folder: {subfolder}")
                print(f"Files and folders in {subfolder}:")

                # Recursive search within "Archive" folders (excluding GDB folders)
                for root, dirs, files in os.walk(folder_path):
                    for dir_name in dirs:
                        dir_path = os.path.join(root, dir_name)
                        try:
                            modification_date = datetime.fromtimestamp(os.path.getmtime(dir_path))
                        except FileNotFoundError:
                            modification_date = None  # Handle the case where the folder is not found
                        print(f"  - {dir_name} (Date Modified: {modification_date})")

                        # Write the folder details to the Excel
                        output_worksheet.append([foldername, subfolder, modification_date, dir_path])

                    for file_name in files:
                        file_path = os.path.join(root, file_name)
                        try:
                            modification_date = datetime.fromtimestamp(os.path.getmtime(file_path))
                        except FileNotFoundError:
                            modification_date = None  # Handle the case where the file is not found
                        print(f"  - {file_name} (Date Modified: {modification_date})")

                        # Write the file details to the Excel
                        output_worksheet.append([foldername, subfolder, modification_date, file_path])

    return archive_folders

# Prompt the user to input the directory path
root_directory = input("Enter the directory path to start searching: ").strip()

# Check if the specified directory exists
if not os.path.isdir(root_directory):
    print("The specified directory does not exist.")
else:
    # Get the user's Downloads folder path
    downloads_folder = os.path.expanduser("~/Downloads")

    # Create the Excel file path in the Downloads folder
    excel_filename = os.path.join(downloads_folder, "archive_folders_with_files.xlsx")

    # Create an Excel workbook
    workbook = Workbook()
    worksheet = workbook.active

    # Write the headers and set them in bold
    worksheet.append(["Parent Directory", "Archive Folder", "Date Modified", "File or Folder Path"])
    for cell in worksheet["1:1"]:
        cell.font = Font(bold=True)

    # Find and catalog "Archive" folders
    archive_folders = find_archive_folders(root_directory, worksheet)

    # Save the workbook
    workbook.save(excel_filename)

    print(f"Cataloged {len(archive_folders)} 'Archive' folders (excluding GDB folders) with files and folders in '{excel_filename}'")
