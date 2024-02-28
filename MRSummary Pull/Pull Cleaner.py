import os
import openpyxl

directory = r'C:\Users\mmccarthy\Desktop\Automation Team\Lightsett\MRSummary Pull\MakeReadySummaryAttributes'

# Function to remove a single column by header name
def remove_column(sheet, col_header):
    column_index = openpyxl.utils.column_index_from_string(col_header)
    sheet.delete_cols(column_index)
    print(f"Column '{col_header}' removed.")

# Function to move data from column H to column A
def move_column_H_to_A(sheet):
    max_row = sheet.max_row
    for row in range(1, max_row + 1):
        sheet[f'A{row}'].value = sheet[f'H{row}'].value

# Function to remove column H from the sheet
def remove_column_H(sheet):
    sheet.delete_cols(openpyxl.utils.column_index_from_string('H'))
    print("Column 'H' removed.")

# List all files in the directory
files = os.listdir(directory)

for file in files:
    if file.endswith('.xlsx'):  # Consider only Excel files
        file_path = os.path.join(directory, file)

        # Load the Excel file
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        # Check if row 2 has no data (except for empty cells)
        row_2_values = sheet[2]
        empty_row_2 = all(cell.value in (None, '') for cell in row_2_values)

        # If row 2 is empty, delete the file and continue to the next file
        if empty_row_2:
            os.remove(file_path)
            print(f"File '{file}' deleted as row 2 contains no data.")
            continue

        # Remove specified columns one at a time
        columns_to_remove = ['R', 'Q', 'P', 'O', 'N', 'L', 'K', 'J', 'I']
        for col in columns_to_remove:
            remove_column(sheet, col)

        # Move data from column H to column A
        move_column_H_to_A(sheet)

        # Remove column H after moving its data to column A
        remove_column_H(sheet)

        # Save the modified workbook
        workbook.save(file_path)
        print(f"Data moved from column H to column A, and column H completely removed in file '{file}'.")

print("Process completed.")
