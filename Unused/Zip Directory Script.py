import os
import zipfile
from openpyxl import Workbook

def extract_folders_from_zip(zip_file):
    folder_list = set()
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        for file_name in zip_ref.namelist():
            # Extract folder names from file paths
            folder_name = os.path.dirname(file_name)
            folder_list.add(folder_name)
    return list(folder_list)

def create_xlsx_from_zip_contents(directory):
    wb = Workbook()
    ws = wb.active
    ws.title = "Zipped Contents"

    # Get a list of all zip files in the directory
    zip_files = [file for file in os.listdir(directory) if file.endswith(".zip")]

    current_row = 1
    current_col = 1
    level = 0

    def write_folder(folder_name, row, col):
        ws.cell(row=row, column=col, value=folder_name)

    for zip_file in zip_files:
        folder_list = extract_folders_from_zip(os.path.join(directory, zip_file))
        folder_name = os.path.splitext(zip_file)[0]

        # Write folder name at the current level
        write_folder(folder_name, current_row, current_col)
        current_row += 1

        # Increase the level
        level += 1

        for folder_name in folder_list:
            # Indent subfolders based on the level
            write_folder(folder_name, current_row, current_col + level)
            current_row += 1

        # Reset the level for the next folder
        level = 0

    output_file = os.path.join(directory, "File_Glossary.xlsx")
    wb.save(output_file)
    print(f"Excel file with folder names from zipped contents saved as '{output_file}'")

if __name__ == "__main__":
    directory_path = r"X:\Customers\23 V1Fiber\23.065.001 V1 Corning Lima - Google Drive Backup\20 November 2023"  # Replace with the path to your directory containing the zip files
    create_xlsx_from_zip_contents(directory_path)
