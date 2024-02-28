import arcpy
import os
import shutil
from datetime import datetime

def export_map_to_kmz(aprx_path, kmz_paths):
    if not aprx_path.lower().endswith('.aprx'):
        raise ValueError("The provided file is not an ArcGIS Pro project file (.aprx)")

    aprx = arcpy.mp.ArcGISProject(aprx_path)

    kmz_map = None
    for map in aprx.listMaps():
        if map.name == "KMZ Map":
            kmz_map = map
            break

    if not kmz_map:
        raise ValueError("No map named 'KMZ Map' found in the project.")

    parent_directory = os.path.dirname(aprx_path)
    upper_directory = os.path.dirname(parent_directory)
    kmz_folder = os.path.join(upper_directory, "KMZ")

    if not os.path.exists(kmz_folder):
        os.makedirs(kmz_folder)

    parent_folder_name = os.path.basename(parent_directory)
    output_kmz = os.path.join(kmz_folder, f"{parent_folder_name}.kmz")

    # Check if the KMZ file already exists
    if os.path.exists(output_kmz):
        # Append the current date and time to the file name
        date_str = datetime.now().strftime("%H_%M_%y_%m_%d")
        file_name, file_extension = os.path.splitext(output_kmz)
        output_kmz = f"{file_name}_{date_str}{file_extension}"
        print(f"An existing KMZ with the same name found. Renaming to {output_kmz}")

    # Corrected MapToKML call
    arcpy.conversion.MapToKML(kmz_map, output_kmz)

    print(f"Exported KMZ file to {output_kmz}")
    
    # Append the KMZ file path to the list
    kmz_paths.append(output_kmz)
    
    # Check if the "_Archive" folder exists, and create it if necessary
    archive_folder = os.path.join(upper_directory, "_Archive")
    if not os.path.exists(archive_folder):
        os.makedirs(archive_folder)
    
    # Move the renamed KMZ file to the "_Archive" folder
    archived_kmz = os.path.join(archive_folder, os.path.basename(output_kmz))
    shutil.move(output_kmz, archived_kmz)
    print(f"Archived KMZ to {archived_kmz}")

def process_aprx_paths(file_path):
    kmz_paths = []  # List to store KMZ file paths
    
    with open(file_path, 'r') as file:
        for line in file:
            aprx_path = line.strip()
            if aprx_path:
                try:
                    export_map_to_kmz(aprx_path, kmz_paths)
                except Exception as e:
                    print(f"Error processing {aprx_path}: {e}")
    
    # Write the KMZ file paths to "KMZ_Paths.txt"
    script_directory = os.path.dirname(os.path.abspath(__file__))
    kmz_paths_file = os.path.join(script_directory, "KMZ_Paths.txt")
    
    with open(kmz_paths_file, 'w') as kmz_file:
        for path in kmz_paths:
            kmz_file.write(path + '\n')

# Example usage
paths_file = r"C:\Users\mmccarthy\Desktop\Automation Team\_real_aprx_without_kmz_map_log.txt"
process_aprx_paths(paths_file)
