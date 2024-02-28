import arcpy
import os

def import_mapx_and_repath(target_aprx_path, mapx_path, gdb_name="lld_design.gdb"):
    try:
        # Extract the project name from the target APRX file path
        project_name = os.path.splitext(os.path.basename(target_aprx_path))[0]

        print(f"Processing target APRX file '{target_aprx_path}'")

        # Load the target APRX project
        aprx = arcpy.mp.ArcGISProject(target_aprx_path)
        print("Target APRX file loaded successfully.")

        # Check if "KMZ Map" exists
        kmz_map = next((m for m in aprx.listMaps() if m.name == "KMZ Map"), None)

        if not kmz_map:
            # Import the MAPX file if "KMZ Map" does not exist
            aprx.importDocument(mapx_path)
            print(f"Imported MAPX file into '{target_aprx_path}'")
            # Re-fetch the "KMZ Map"
            kmz_map = next((m for m in aprx.listMaps() if m.name == "KMZ Map"), None)

        # Update data sources to the project's lld_design.gdb
        gdb_path = os.path.join(os.path.dirname(target_aprx_path), gdb_name)
        if not os.path.exists(gdb_path):
            raise FileNotFoundError(f"lld_design.gdb not found at: {gdb_path}")

        for lyr in kmz_map.listLayers():
            if lyr.supports("DATASOURCE") and lyr.connectionProperties:
                old_conn_props = lyr.connectionProperties
                if old_conn_props:
                    new_conn_props = old_conn_props.copy()
                    new_conn_props['connection_info']['database'] = gdb_path
                    lyr.updateConnectionProperties(old_conn_props, new_conn_props)

            # Rename "Evans Zone 22" group layer to the project name
            for group_layer in kmz_map.listLayers():
                if group_layer.isGroupLayer and group_layer.name == "Evans Zone 22":
                    group_layer.name = project_name

        aprx.save()
        print(f"Updated 'KMZ Map' and saved '{target_aprx_path}'")

    except Exception as e:
        print(f"General error occurred: {type(e).__name__}: {e}")

def validate_group_layer_name(target_aprx_path, expected_group_name):
    try:
        aprx = arcpy.mp.ArcGISProject(target_aprx_path)
        kmz_map = next((m for m in aprx.listMaps() if m.name == "KMZ Map"), None)
        
        if not kmz_map:
            print(f"No 'KMZ Map' found in '{target_aprx_path}'.")
            return False

        group_layer = next((lyr for lyr in kmz_map.listLayers() if lyr.isGroupLayer and lyr.name == expected_group_name), None)
        
        if group_layer:
            print(f"Group layer '{expected_group_name}' found in 'KMZ Map' of '{target_aprx_path}'.")
            return True
        else:
            print(f"No group layer named '{expected_group_name}' found in 'KMZ Map' of '{target_aprx_path}'.")
            return False

    except Exception as e:
        print(f"Error occurred while validating group layer name: {type(e).__name__}: {e}")
        return False

def read_target_aprx_paths(file_path):
    with open(file_path, 'r') as file:
        return [line.strip() for line in file]

def user_confirmation(file_paths):
    print("List of APRX file paths to be processed:")
    for path in file_paths:
        print(path)
    response = input("Is this list correct? (yes/no): ").strip().lower()
    return response in ['yes', 'y']

if __name__ == "__main__":
    mapx_path = r"C:\Users\mmccarthy\Desktop\Automation Team\KMZ Map.mapx"
    aprx_file_list_path = r"C:\Users\mmccarthy\Desktop\Automation Team\_real_aprx_without_kmz_map_log.txt"
    target_aprx_paths = read_target_aprx_paths(aprx_file_list_path)

    if user_confirmation(target_aprx_paths):
        for target_aprx_path in target_aprx_paths:
            project_name = os.path.splitext(os.path.basename(target_aprx_path))[0]
            import_mapx_and_repath(target_aprx_path, mapx_path)
            validate_group_layer_name(target_aprx_path, project_name)
        print("Script completed.")
    else:
        print("Script exited. The list was not confirmed.")

