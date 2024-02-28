import os
import zipfile
import xml.etree.ElementTree as ET

def rename_kml_map_layer(kmz_file):
    # Extract the KMZ file to a temporary directory
    with zipfile.ZipFile(kmz_file, 'r') as zip_ref:
        temp_dir = 'temp_kmz'
        zip_ref.extractall(temp_dir)

    # Find and rename the KML Map layer
    kml_file = os.path.join(temp_dir, 'doc.kml')
    tree = ET.parse(kml_file)
    root = tree.getroot()

    # Define KML namespace
    kml_namespace = {'kml': 'http://www.opengis.net/kml/2.2'}

    # Register the namespace for XPath queries
    ET.register_namespace('', kml_namespace['kml'])

    for placemark in root.findall('.//kml:Placemark', namespaces=kml_namespace):
        name_element = placemark.find('.//kml:name', namespaces=kml_namespace)
        if name_element is not None and name_element.text == 'KML Map':
            name_element.text = os.path.splitext(os.path.basename(kmz_file))[0]

    tree.write(kml_file)

    # Remove empty KML files
    for kml_file in os.listdir(temp_dir):
        if kml_file.endswith('.kml'):
            kml_path = os.path.join(temp_dir, kml_file)
            if os.path.getsize(kml_path) == 0:
                os.remove(kml_path)

    # Rename the folder underneath KML Map to "Design_Features"
    for folder in root.findall('.//kml:Folder', namespaces=kml_namespace):
        name_element = folder.find('.//kml:name', namespaces=kml_namespace)
        if name_element is not None and name_element.text == 'KML Map':
            name_element.text = 'Design_Features'

    # Create a new KMZ file with the cleaned-up data
    cleaned_kmz_file = os.path.splitext(kmz_file)[0] + '_cleaned.kmz'
    with zipfile.ZipFile(cleaned_kmz_file, 'w', zipfile.ZIP_DEFLATED) as cleaned_zip:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                cleaned_zip.write(file_path, os.path.relpath(file_path, temp_dir))

    # Clean up temporary files and directory
    os.remove(kmz_file)
    os.rename(cleaned_kmz_file, kmz_file)
    for file in os.listdir(temp_dir):
        file_path = os.path.join(temp_dir, file)
        os.remove(file_path)
    os.rmdir(temp_dir)

if __name__ == "__main__":
    kmz_directory = r'S:\Customers\48 Glass Roots\48.001.001 Evans, CO\Evans, CO\2. LLD\FDA 1.12A\KMZ'

    for root, dirs, files in os.walk(kmz_directory):
        for file in files:
            if file.endswith('.kmz'):
                kmz_file_path = os.path.join(root, file)
                rename_kml_map_layer(kmz_file_path)
