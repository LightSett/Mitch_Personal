import arcpy
import os
import csv
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog

def export_attribute_rules(template_path, feature_classes):

    #Set the current workspace to that gdb
        arcpy.env.overwriteOutput = True
        arcpy.env.workspace = template_path

        current_directory = os.getcwd()
        sub_directory = 'ar_csv'


        # Loop through each feature class
        for feature_class_name in feature_classes:
            feature_class_path = os.path.join(template_path, feature_class_name)
            
            try:
                # Check if the feature class exists
                if arcpy.Exists(feature_class_path):
                    # Export attribute rules to a CSV file
                    out_csv_file = os.path.join(current_directory, sub_directory, feature_classes[feature_class_name])
                    arcpy.ExportAttributeRules_management(feature_class_path, out_csv_file)
                    print(f"Attribute rules exported for '{feature_class_name}' into '{out_csv_file}'.")

            except arcpy.ExecuteError as e:
                print("An error occurred:", e)

def delete_attribute_rules(gdb_paths, feature_classes):

    user_home = os.path.expanduser("~")
    downloads_folder = None
    
    if os.name == 'posix':  # Linux or macOS
        downloads_folder = os.path.join(user_home, 'Downloads')
    elif os.name == 'nt':   # Windows
        downloads_folder = os.path.join(user_home, 'Downloads')    

    #A loop to iterate through each LLD.gdb that was selected.
    for gdb_path in gdb_paths: 

        # Set the workspace (geodatabase) where the feature classes are located
        gdb_path = gdb_path
        export_path = downloads_folder

        #Set the current workspace to that gdb
        arcpy.env.overwriteOutput = True
        arcpy.env.workspace = gdb_path

        print("working")

        # Get a list of all the different feature classes.
        all_feature_classes = arcpy.ListFeatureClasses()

        print(all_feature_classes)

        # Loop through each feature class
        for feature_class_name in feature_classes:
            feature_class_path = os.path.join(gdb_path, feature_class_name)
            rule_names = []
            
            try:
                # Check if the feature class exists
                if arcpy.Exists(feature_class_path):
                    # Export attribute rules to a CSV file
                    out_csv_file = export_path + "/" + feature_class_name +".csv"
                    out_xml_file = export_path + "/" + feature_class_name + ".csv.xml"
                    arcpy.ExportAttributeRules_management(feature_class_path, out_csv_file)
                    print(f"Attribute rules exported for '{feature_class_name}' to '{out_csv_file}'.")
                    
                    # Read the CSV and delete attribute rules
                    with open(out_csv_file, 'r') as csv_file:
                        csv_reader = csv.reader(csv_file)
                        next(csv_reader)  # Skip the header row
                        for row in csv_reader:
                            rule_names.append(row[0])
                            
                        arcpy.management.DeleteAttributeRule(feature_class_path, rule_names)
                        print(f"Deleted attribute rules '{rule_names}' from '{feature_class_name}'.")
                    
                    # Delete the CSV file
                    os.remove(out_csv_file)
                    os.remove(out_xml_file)
                    print(f"CSV file '{out_csv_file}' deleted.")
                else:
                    print(f"Feature class '{feature_class_name}' does not exist.")
            except arcpy.ExecuteError as e:
                print("An error occurred:", e)

def import_attribute_rules(gdb_paths, feature_classes):   

    #A loop to iterate through each LLD.gdb that was selected.
    for gdb_path in gdb_paths: 

        # Set the workspace (geodatabase) where the feature classes are located
        gdb_path = gdb_path
        current_directory = os.getcwd()
        sub_directory = 'ar_csv'

        #Set the current workspace to that gdb
        arcpy.env.overwriteOutput = True
        arcpy.env.workspace = gdb_path

        print("working")

        # Get a list of all the different feature classes.
        all_feature_classes = arcpy.ListFeatureClasses()

        print(all_feature_classes)

        # Loop through each feature class
        for feature_class_name in feature_classes:
            feature_class_path = os.path.join(gdb_path, feature_class_name)
            
            
            try:
                # Check if the feature class exists
                if arcpy.Exists(feature_class_path):
                    # import attribute rules from a CSV file
                    in_csv_file = os.path.join(current_directory, sub_directory, feature_classes[feature_class_name])
                    print(feature_classes[feature_class_name])
                    arcpy.management.ImportAttributeRules(feature_class_path, in_csv_file)
                    print(f"Attribute rules imported for '{feature_class_name}' from '{in_csv_file}'.")
        
                else:
                    print(f"Feature class '{feature_class_name}' does not exist.")
            except arcpy.ExecuteError as e:
                print("An error occurred:", e)




feature_classes = {
    "Demand_Point": 'demand_point_ar.CSV',
    "Equipment": 'equipment_ar.csv',
    "Slack_Loop": 'slack_loop_ar.csv',
    "Splice_Closure": 'splice_closure_ar.csv',
    "Pole": 'pole_ar.csv',
    "UG_Structure": 'ug_structure_ar.csv',
    "Strand": 'strand_ar.csv',
    "Conduit": 'conduit_ar.csv',
    "Transmedia": 'transmedia_ar.csv',
    "Span": 'span_ar.csv',
    "Serving_Area": 'serving_area_ar.csv',
    "FDA": 'fda_ar.csv',
    "Design_Tools": 'design_tools_ar.csv',
    "Project_Details": 'project_details_ar.csv',
    "IQGeo_Tools": 'iqgeo_tools_ar.csv',
    "RoutingHelper": 'routinghelper_ar.csv',
    "BOM": 'bom_ar.csv',
    "AssignmentHelper": 'assignmenthelper_ar.csv',
}

template_path = r"S:\Customers\23 V1Fiber\23.035.001 V1 Cloudwyze\Nashville Extension (Phase 5)\GIS\LLD\5.1\Asbuilt\V1_Cloudwyze_5_1_LLD\lld_design.gdb"
gdb_paths = []

def load_gdb_file_paths():
    global gdb_paths
    #Open the file dialog to ask the user for a LLD Path
    LLD_path = filedialog.askdirectory(title="Select a LLD folder")
    #Append the selected LLD.gdb to the gdb_paths variable
    gdb_paths.append(LLD_path)
    
def run_script():
    global template_path, feature_classes, gdb_paths
    export_attribute_rules(template_path, feature_classes)
    delete_attribute_rules(gdb_paths, feature_classes)
    import_attribute_rules(gdb_paths, feature_classes)
    messagebox.showinfo("Script Completed", "Attribute rules updated!")

def close_app():
    root.destroy()

root = tk.Tk()
root.title("Attribute Rules Management")

# GUI elements
filepath_label = tk.Label(root, text="Navigate to the LLD_Design.gdb to be updated:")
filepath_label.pack()

filepath_entry = tk.Button(root, text="Load GDB", command=load_gdb_file_paths)
filepath_entry.pack()

run_button = tk.Button(root, text="Run Script", command=run_script)
run_button.pack()

close_button = tk.Button(root, text="Close", command=close_app)
close_button.pack()


def browse_file():
    filepath = filedialog.askopenfilename()
    filepath_entry.delete(0, tk.END)
    filepath_entry.insert(0, filepath)

## Replace the gdb paths list with the file paths to your lld.gdbs.
## Can be multiple lld.gdb's, just separate them using a ,
#gdb_paths = [
#    "S:/Customers/23 V1Fiber/23.065.001 V1 Corning Lima/Scranton/LLD/Route 1/1105PB/1105PB Revise/1105PB/LLD_Design.gdb",
#]


#delete_attribute_rules(gdb_paths, feature_classes)
#import_attribute_rules(gdb_paths, feature_classes)