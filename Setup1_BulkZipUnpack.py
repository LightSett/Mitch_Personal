import zipfile
import os

def extract_zip_to_directories(zip_path, directories):
    # Ensure the ZIP file exists
    if not os.path.exists(zip_path):
        print(f"ZIP file does not exist: {zip_path}")
        return

    for directory in directories:
        # Check if the directory exists
        if not os.path.exists(directory):
            print(f"Directory does not exist, skipping: {directory}")
            continue  # Skip to the next iteration
        
        # Extract the ZIP file
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(directory)
            print(f"Extracted {zip_path} to {directory}")

# Example usage
zip_path = r'C:\Users\mmccarthy\Desktop\Stokes_HLD.zip'  # Update this path to your ZIP file
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
                r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-114\st1_Mapspace',
                r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-115\st1_Mapspace',
                r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-116\st1_Mapspace',
                r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-117\st1_Mapspace',
                r'X:\Customers\53 RiverStreet\53.001.002 Stokes\st1_24-OSPSTOK-118\st1_Mapspace']
  # Update these paths to your target directories
extract_zip_to_directories(zip_path, directories)
