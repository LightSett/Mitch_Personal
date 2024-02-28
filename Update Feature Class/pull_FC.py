import arcpy
import os
import xlwt

# Set the workspace environment
arcpy.env.workspace = r"C:\Users\mmccarthy\Desktop\Test Spaces\FlexNAP_LLD_Template_Mapspace_12-19-23\FlexNAP_LLD_Template_Mapspace\LLD_Design.gdb"  # Change this to your geodatabase path

# Get a list of all feature classes and feature datasets
feature_classes = []
feature_datasets = []

for dirpath, dirnames, filenames in arcpy.da.Walk(arcpy.env.workspace, datatype="FeatureClass", type="ALL"):
    for filename in filenames:
        feature_classes.append(os.path.join(dirpath, filename))

for dirpath, dirnames, filenames in arcpy.da.Walk(arcpy.env.workspace, datatype="FeatureDataset"):
    for dirname in dirnames:
        feature_datasets.append(os.path.join(dirpath, dirname))

# Create an Excel workbook
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("Feature Classes")
sheet.write(0, 0, "Feature Classes")
sheet.write(0, 1, "Feature Datasets")

# Write feature classes and feature datasets to Excel
for i, fc in enumerate(feature_classes):
    sheet.write(i + 1, 0, fc)

for i, fd in enumerate(feature_datasets):
    sheet.write(i + 1, 1, fd)

# Save the Excel file in the same directory as the script
script_directory = os.path.dirname(os.path.abspath(__file__))
excel_file_path = os.path.join(script_directory, "FeatureClassesAndDatasets.xls")
workbook.save(excel_file_path)

print(f"Excel file saved at: {excel_file_path}")
