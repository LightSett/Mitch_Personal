import os
import zipfile
import shutil

# Function to rename the main folder in a KMZ file
def rename_main_folder(kmz_path, new_path):
    with zipfile.ZipFile(kmz_path, 'r') as kmz_file:
        # Extract the contents of the KMZ file to a temporary directory
        temp_dir = 'temp'
        kmz_file.extractall(temp_dir)

        # Get the list of extracted files and folders
        extracted_files = os.listdir(temp_dir)

        if len(extracted_files) == 1 and os.path.isdir(os.path.join(temp_dir, extracted_files[0])):
            main_folder = extracted_files[0]
            # Rename the main folder to match the KMZ file name
            new_folder_name = os.path.splitext(os.path.basename(kmz_path))[0]
            new_folder_path = os.path.join(temp_dir, new_folder_name)
            os.rename(os.path.join(temp_dir, main_folder), new_folder_path)

            # Update the contents of the KMZ file with the renamed folder
            with zipfile.ZipFile(kmz_path, 'w', zipfile.ZIP_DEFLATED) as new_kmz_file:
                for root, _, files in os.walk(new_folder_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, new_folder_path)
                        new_kmz_file.write(file_path, arcname=arcname)

    # After extracting and modifying, rename the KMZ file
    os.rename(kmz_path, new_path)

# Function to remove empty folders recursively
def remove_empty_folders(folder_path):
    for root, dirs, files in os.walk(folder_path, topdown=False):
        for dir in dirs:
            dir_path = os.path.join(root, dir)
            if not os.listdir(dir_path):
                os.rmdir(dir_path)

# Read the list of KMZ file paths from the specified .txt document
kmz_file_paths = []
with open(r'C:\Users\mmccarthy\Desktop\Automation Team\KMZ_Paths.txt', 'r') as file:
    kmz_file_paths = [line.strip() for line in file]

for kmz_path in kmz_file_paths:
    # Check if the file exists at the specified path
    if os.path.exists(kmz_path):
        # Check if the file name already starts with "Evans, CO FDA "
        base_name = os.path.splitext(os.path.basename(kmz_path))[0]
        if base_name.startswith("Evans, CO FDA "):
            # If it already has the prefix, continue with the other operations
            rename_main_folder(kmz_path, kmz_path)
        else:
            # Add "Evans, CO FDA" to the front of the file name
            new_name = f"Evans, CO FDA {base_name}.kmz"
            new_path = os.path.join(os.path.dirname(kmz_path), new_name)

            # Rename the main folder inside the KMZ file and the KMZ file itself
            rename_main_folder(kmz_path, new_path)

        # Remove empty folders inside the KMZ file
        remove_empty_folders(new_path)
    else:
        # If the file does not exist, check the parent folder for "Evans, CO FDA {kmz from list}"
        parent_folder = os.path.dirname(kmz_path)
        expected_kmz_name = f"Evans, CO FDA {os.path.basename(kmz_path)}"
        expected_kmz_path = os.path.join(parent_folder, expected_kmz_name)
        if os.path.exists(expected_kmz_path):
            # Continue with the operations on the expected KMZ file
            rename_main_folder(expected_kmz_path, expected_kmz_path)
            remove_empty_folders(expected_kmz_path)

# Clean up the temporary directory
if os.path.exists('temp'):
    shutil.rmtree('temp')  # Use shutil.rmtree to remove non-empty directories
