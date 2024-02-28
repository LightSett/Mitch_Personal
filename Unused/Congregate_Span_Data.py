import xlwings as xw

# Define the path to the workbook named "Consolidated_Att_Tables"
workbook_path = r"S:\Customers\23 V1Fiber\23.065.001 V1 Corning Lima\Scranton\Weekly Lima Report Data\Cable Length Consolidation\Consolidated_Att_Tables.xlsx"

# Open the workbook
consolidated_wb = xw.Book(workbook_path)

# Create a new sheet at the beginning of the workbook named "all_spans"
new_sheet = consolidated_wb.sheets.add("all_spans")

# Loop through each sheet in the workbook
for sheet in consolidated_wb.sheets:
    if sheet.name != "all_spans":
        # Extract data from columns 4, 7, and 11
        data_to_copy = sheet.range((1, 4), (sheet.api.UsedRange.Rows.count, 11))

        # Determine where to paste the data in the new sheet
        last_row = new_sheet.api.UsedRange.Rows.count
        new_row = last_row + 1 if last_row > 1 else 1

        # Copy the data into columns 1, 2, and 3 of the new sheet
        data_to_copy.copy(destination=new_sheet.range((new_row, 1)))

# Save the updated workbook
consolidated_wb.save()
consolidated_wb.close()

print("A new sheet 'all_spans' has been added to 'Consolidated_Att_Tables.xlsx' with data from columns 4, 7, and 11 of other sheets.")
