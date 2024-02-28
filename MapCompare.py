import os
import arcpy
import PySimpleGUI as sg

def find_design_gdbs(directory):
    design_gdbs = []
    for root, dirs, files in os.walk(directory):
        if not any(folder_name.lower() in root.lower() for folder_name in ["archive", "backup"]):
            for file in files:
                if file.endswith('.aprx'):
                    design_gdbs.append(os.path.join(root, file))
    return design_gdbs

def copy_kmz_map(source_aprx, target_aprx):
    try:
        source_project = arcpy.mp.ArcGISProject(source_aprx)
        source_map = next((map for map in source_project.listMaps() if "KMZ" in map.name), None)

        if source_map:
            target_project = arcpy.mp.ArcGISProject(target_aprx)
            target_map_names = [map.name for map in target_project.listMaps()]

            if "KMZ" not in target_map_names:
                print(f"Copying 'KMZ' map from '{source_aprx}' to '{target_aprx}'...")
                target_project.importDocument(source_aprx)
                print("Map copied successfully.")
            else:
                print(f"'KMZ' map already exists in '{target_aprx}'.")

    except Exception as e:
        print(f"Error copying 'KMZ' map: {e}")

import logging

# Set up logging
logging.basicConfig(filename='update_aprx.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def copy_kmz_map(source_aprx, target_aprx):
    try:
        source_project = arcpy.mp.ArcGISProject(source_aprx)
        source_map = next((map for map in source_project.listMaps() if "KMZ" in map.name), None)

        if source_map:
            target_project = arcpy.mp.ArcGISProject(target_aprx)
            target_map_names = [map.name for map in target_project.listMaps()]

            if "KMZ" not in target_map_names:
                logging.info(f"Copying 'KMZ' map from '{source_aprx}' to '{target_aprx}'...")
                target_project.importDocument(source_aprx)
                logging.info("Map copied successfully.")
            else:
                logging.info(f"'KMZ' map already exists in '{target_aprx}'.")

    except Exception as e:
        logging.error(f"Error copying 'KMZ' map: {e}")

def update_data_sources(source_aprx, target_aprx):
    try:
        source_project = arcpy.mp.ArcGISProject(source_aprx)
        target_project = arcpy.mp.ArcGISProject(target_aprx)

        source_maps = source_project.listMaps()
        target_maps = target_project.listMaps()

        # Keep track of processed layers to detect infinite loops
        processed_layers = set()

        total_layers = sum(len(source_map.listLayers()) for source_map in source_maps)

        for source_map in source_maps:
            for source_layer in source_map.listLayers():
                for target_map in target_maps:
                    target_layer = next((layer for layer in target_map.listLayers() if layer.name == source_layer.name), None)
                    if target_layer:
                        logging.info(f"Updating data source for {target_layer.name} in '{target_map.name}'...")
                        target_layer.updateConnectionProperties(source_layer.connectionProperties)
                        logging.info(f"Data source updated for {target_layer.name} in '{target_map.name}'.")
                        
                        # Check for infinite loop
                        layer_key = f"{source_map.name}_{source_layer.name}"
                        if layer_key in processed_layers:
                            logging.warning(f"Infinite loop detected for layer {source_layer.name} in map {source_map.name}.")
                        else:
                            processed_layers.add(layer_key)
                    
                    logging.info(f"Progress: {len(processed_layers)}/{total_layers} layers updated ({(len(processed_layers)/total_layers)*100:.2f}%).")

    except Exception as e:
        logging.error(f"Error updating data sources: {e}")

def main():
    source_aprx = sg.popup_get_file('Select Source APRX', file_types=(("ArcGIS Project Files", "*.aprx"),))
    if not source_aprx:
        return

    directory = sg.popup_get_folder('Select Directory Containing Target APRX Files')
    if not directory:
        return

    target_aprx_list = find_design_gdbs(directory)

    if not target_aprx_list:
        sg.popup('No target APRX files found in the specified directory.')
        return

    layout = [[sg.Checkbox(os.path.basename(target_aprx), key=target_aprx)] for target_aprx in target_aprx_list]
    layout += [[sg.Button('Update Selected'), sg.Button('Cancel')]]

    window = sg.Window('APRX Selection', layout)

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Cancel'):
            break
        if event == 'Update Selected':
            update_warning = sg.popup_yes_no('WARNING: This action may create an infinite loop in your project. Continue?')
            if update_warning == 'Yes':
                for target_aprx in target_aprx_list:
                    if values[target_aprx]:
                        print(f"Updating '{target_aprx}' with data from '{source_aprx}':")
                        copy_kmz_map(source_aprx, target_aprx)
                        update_data_sources(source_aprx, target_aprx)
                sg.popup('Update Complete')
            else:
                print("Update operation canceled.")
            break

    window.close()

if __name__ == "__main__":
    main()
