import os

def rename_aprx_files(directories):
    for directory in directories:
        # Ensure the directory exists
        if not os.path.isdir(directory):
            print(f"The directory does not exist: {directory}")
            continue

        # Search for .aprx files in the directory
        for item in os.listdir(directory):
            if item.endswith('.aprx'):
                aprx_file_path = os.path.join(directory, item)
                parent_directory_name = os.path.basename(os.path.dirname(directory))
                new_aprx_file_name = f"{parent_directory_name}.aprx"
                new_aprx_file_path = os.path.join(directory, new_aprx_file_name)

                # Rename the .aprx file
                os.rename(aprx_file_path, new_aprx_file_path)
                print(f"Renamed {aprx_file_path} to {new_aprx_file_path}")

directories = [r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_22-OSPSTOK-113\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_22-OSPSTOK-114\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_23-OSPSTOK-103\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_23-OSPSTOK-104\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-102\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-103\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-104\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-105\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-106\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-107\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-108\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-109\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-110\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-111\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-112\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-113\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-115\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-116\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-117\st1_Mapspace',
            r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-118\st1_Mapspace']


rename_aprx_files(directories)
