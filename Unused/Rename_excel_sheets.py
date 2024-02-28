import os
import xlwings as xw

# Define the folder path containing the Excel files
folder_path = r"S:\Customers\23 V1Fiber\23.065.001 V1 Corning Lima\Scranton\Weekly Lima Report Data\Cable Length Consolidation"

# Get a list of Excel files in the folder
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

# Create a new Excel workbook for consolidation
consolidated_wb = xw.Book()
consolidated_wb.save(r"S:\Customers\23 V1Fiber\23.065.001 V1 Corning Lima\Scranton\Weekly Lima Report Data\Cable Length Consolidation\Consolidated_Att_Tables.xlsx")

# Loop through each Excel file
for excel_file in excel_files:
    file_path = os.path.join(folder_path, excel_file)

    # Extract the first 6 characters of the workbook name
    new_sheet_name_prefix = excel_file[:6]

    # Open the Excel file
    wb = xw.Book(file_path)

    print(f"Modifying workbook: {wb.name}")

    # Loop through each sheet in the workbook
    for sheet in wb.sheets:
        if "transmedia" in sheet.name.lower():  # Check if "transmedia" is in the sheet name (case-insensitive)
            sheet.name = f"{new_sheet_name_prefix}_{sheet.name}"
            sheet.api.Copy(Before=consolidated_wb.sheets[0].api)

    wb.close()

# Save the consolidated workbook
consolidated_wb.save(r"S:\Customers\23 V1Fiber\23.065.001 V1 Corning Lima\Scranton\Weekly Lima Report Data\Cable Length Consolidation\Consolidated_Att_Tables.xlsx")
consolidated_wb.close()

print("Sheets with 'span' in their names (case-insensitive) have been consolidated into 'Consolidated_Att_Tables.xlsx'.")
