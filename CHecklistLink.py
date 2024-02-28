import openpyxl
import os

def link_checkboxes_to_column_g(excel_file):
    # Load the workbook and get the active sheet
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active

    # Assuming checkboxes are linked to cells, find all cells with links
    # and create corresponding links in column G
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == 'TRUE' or cell.value == 'FALSE':
                # Link this cell to the equivalent row in column G
                sheet[f'G{cell.row}'] = cell.value

    # Prepare the new filename
    directory, filename = os.path.split(excel_file)
    new_filename = os.path.join(directory, 'modified_' + filename)

    # Save the modified workbook
    workbook.save(new_filename)

# Example usage
link_checkboxes_to_column_g('C:\\Users\\mmccarthy\\Downloads\\Brush Checklist.xlsx')
