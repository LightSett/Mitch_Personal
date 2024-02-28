import openpyxl
import os
import csv

def convert_xlsx_to_csv(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx') and not file_name.startswith('~$'):
            try:
                xlsx_file_path = os.path.join(folder_path, file_name)
                wb = openpyxl.load_workbook(xlsx_file_path)
                sheet = wb.active

                csv_file_path = os.path.join(folder_path, file_name.replace('.xlsx', '.csv'))
                with open(csv_file_path, 'w', newline='', encoding='utf-8') as csv_file:
                    writer = csv.writer(csv_file)
                    for row in sheet.iter_rows():
                        writer.writerow([cell.value for cell in row])
                    
                # Rename the CSV file to remove 'XLSX' or '_XLSX' from the name
                new_csv_file_path = csv_file_path.replace(' XLSX', '_CSV').replace('_XLSX', '_CSV')  # Remove 'XLSX' or '_XLSX' from the filename
                os.rename(csv_file_path, new_csv_file_path)
            except Exception as e:
                print(f"Error converting {file_name} to CSV: {str(e)}")

# Replace 'folder_path' with the actual folder path containing your Excel files
folder_path = r"C:\Users\mmccarthy\Desktop\1101PA Make Ready Info\XLSX"
convert_xlsx_to_csv(folder_path)
