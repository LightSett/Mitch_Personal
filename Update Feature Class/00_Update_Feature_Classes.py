import arcpy
import tkinter as tk
from tkinter import filedialog, messagebox

def update_fields(source_gdb, target_gdb):
    def disable_attribute_rules(feature_classes):
        for fc in feature_classes:
            desc = arcpy.Describe(fc)
            if desc.HasAttributeRules:
                try:
                    arcpy.DisableAttributeRules_management(fc)
                    print(f"Attribute rules disabled for {fc}")
                except arcpy.ExecuteError as e:
                    print(f"Error while disabling attribute rules for {fc}: {e}")
            else:
                print(f"No attribute rules found for {fc}. Skipping disable operation.")

    def enable_attribute_rules(feature_classes):
        for fc in feature_classes:
            desc = arcpy.Describe(fc)
            if desc.HasAttributeRules:
                try:
                    arcpy.EnableAttributeRules_management(fc)
                    print(f"Attribute rules enabled for {fc}")
                except arcpy.ExecuteError as e:
                    print(f"Error while enabling attribute rules for {fc}: {e}")
            else:
                print(f"No attribute rules found for {fc}. Skipping enable operation.")

    ignored_field_names = ["OBJECTID", "OBJECT_ID", "GlobalID", "GUID", "shape", "Shape", "SHAPE"]
    feature_classes_to_update = [
        "Demand_Point", "Equipment", "Slack_Loop", "Splice_Closure", "Pole", "UG_Structure", "Strand", "Conduit",
        "Transmedia", "Span", "Serving_Area", "FDA", "Design_Tools", "Project_Details", "IQGeo_Tools",
        "RoutingHelper", "AssignmentHelper",
        "Rev_Cloud", "Permit_Polygon", "Permit", "Phase_Boundaries", "Field_Note_Line", "Dimensions",
        "Pole_Proposals", "OffsetHelper", "Graphics_Line", "Graphics_Point", "Detail_Indicator",
        "Graphics_Polygon", "AssignmentHelper", "RoutingHelper", "Demand_PointHelper"
    ]

    def update_fields_operation():
        disable_attribute_rules(feature_classes_to_update)

        arcpy.env.workspace = target_gdb

        for template_feature_class in feature_classes_to_update:
            source_feature_class_path = f"{source_gdb}/{template_feature_class}"
            target_feature_class_path = f"{target_gdb}/{template_feature_class}"

            if arcpy.Exists(source_feature_class_path) and arcpy.Exists(target_feature_class_path):
                template_fields = {f.name: f for f in arcpy.ListFields(source_feature_class_path)}
                target_fields = {f.name: f for f in arcpy.ListFields(target_feature_class_path)}

                for template_field_name, template_field in template_fields.items():
                    if (template_field_name not in ignored_field_names and 
                            not template_field_name.endswith('_new') and 
                            template_field_name not in target_fields):
                        try:
                            arcpy.AddField_management(target_feature_class_path, template_field_name, "TEXT",
                                                    field_length=template_field.length,
                                                    field_alias=template_field.aliasName,
                                                    field_is_nullable=template_field.isNullable,
                                                    field_is_required=template_field.required,
                                                    field_domain=template_field.domain)
                            print(f"Added new field '{template_field_name}' to '{target_feature_class_path}'")

                            arcpy.CalculateField_management(target_feature_class_path, template_field_name, f"!{template_field_name}!", "PYTHON3")
                            print(f"Copied values to the new field '{template_field_name}'")

                            arcpy.DeleteField_management(target_feature_class_path, template_field_name)
                            print(f"Deleted old field '{template_field_name}' from '{target_feature_class_path}'")
                        except arcpy.ExecuteError as e:
                            print(f"Error while updating field '{template_field_name}' in '{target_feature_class_path}': {e}")

        enable_attribute_rules(feature_classes_to_update)

    update_fields_operation()


def browse_and_extract(label_text, entry_widget):
    gdb_path = filedialog.askdirectory()
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, gdb_path)


def run_extraction():
    source_gdb = entry_source.get()
    target_gdb = entry_target.get()

    if not source_gdb or not target_gdb:
        tk.messagebox.showwarning("Warning", "Please select both Source and Target Geodatabases.")
        return

    button_browse_source.config(state=tk.DISABLED)
    button_browse_target.config(state=tk.DISABLED)
    button_run.config(state=tk.DISABLED)
    button_close.config(state=tk.DISABLED)
    entry_source.config(state=tk.DISABLED)
    entry_target.config(state=tk.DISABLED)

    update_fields(source_gdb, target_gdb)


root = tk.Tk()
root.title("Update Fields in Target GDB")

label_source = tk.Label(root, text="Source GDB:")
label_source.grid(row=0, column=0)

entry_source = tk.Entry(root, width=50)
entry_source.grid(row=0, column=1)

button_browse_source = tk.Button(root, text="Browse", command=lambda: browse_and_extract("Source GDB:", entry_source))
button_browse_source.grid(row=0, column=2)

label_target = tk.Label(root, text="Target GDB:")
label_target.grid(row=1, column=0)

entry_target = tk.Entry(root, width=50)
entry_target.grid(row=1, column=1)

button_browse_target = tk.Button(root, text="Browse", command=lambda: browse_and_extract("Target GDB:", entry_target))
button_browse_target.grid(row=1, column=2)

button_run = tk.Button(root, text="Run", command=run_extraction)
button_run.grid(row=2, column=1)

button_close = tk.Button(root, text="Close", command=root.destroy)
button_close.grid(row=2, column=2)

root.mainloop()
