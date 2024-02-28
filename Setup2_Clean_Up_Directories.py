import os
import shutil

def move_contents_to_parent_and_remove(directory):
    # Check if the directory exists
    if not os.path.isdir(directory):
        print(f"The directory does not exist: {directory}")
        return

    parent_directory = os.path.dirname(directory)
    if not os.path.isdir(parent_directory):
        print(f"The parent directory does not exist: {parent_directory}")
        return

    # Move each item in the directory to the parent directory
    for item_name in os.listdir(directory):
        item_path = os.path.join(directory, item_name)
        new_path = os.path.join(parent_directory, item_name)

        # Ensure there's no name conflict in the parent directory
        if os.path.exists(new_path):
            print(f"Conflict: An item named {item_name} already exists in {parent_directory}. Skipping...")
            continue

        shutil.move(item_path, parent_directory)
        print(f"Moved {item_path} to {parent_directory}")

    # Remove the original directory
    try:
        os.rmdir(directory)
        print(f"Removed the original directory: {directory}")
    except OSError as e:
        print(f"Error: Failed to remove the original directory {directory}. It may not be empty. Error: {e}")

# Example usage with multiple directories
directories = [r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_22-OSPSTOK-113\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_22-OSPSTOK-114\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_23-OSPSTOK-103\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_23-OSPSTOK-104\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-102\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-103\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-104\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-105\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-106\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-107\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-108\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-109\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-110\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-111\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-112\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-113\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-115\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-116\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-117\st1_Mapspace\Stokes_HLD',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-118\st1_Mapspace\Stokes_HLD']
  # Update this path to the directory you want to move and delete
for directory in directories:
    move_contents_to_parent_and_remove(directory)
