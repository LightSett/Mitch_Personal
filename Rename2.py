import os

# Function to process a single KMZ file path
def process_kmz_file(kmz_path):
    # Extract the file name from the KMZ path
    kmz_name = os.path.basename(kmz_path)

    # Check if the KMZ path contains "archive" or "Archive," and if so, skip it
    if "archive" in kmz_name.lower():
        return None

    # Extract the base name without extension
    base_name, ext = os.path.splitext(kmz_name)

    # Check if "Evans", ",CO", and "FDA" are already present in order
    if "Evans" in base_name and ",CO" in base_name and "FDA" in base_name:
        return kmz_path  # No changes needed, return the original path

    # Build the desired file name
    new_name = "Evans, CO FDA " + base_name

    # Create the new file path with the desired name
    new_path = os.path.join(os.path.dirname(kmz_path), new_name + ext)

    # Rename the KMZ file with the new name
    os.rename(kmz_path, new_path)

    return new_path  # Return the updated path

# Specify the path to the input file containing the list of KMZ file paths
input_file_path = r'C:\Users\mmccarthy\Downloads\kmz_file_list.txt'

# Read the list of KMZ file paths from the input file
kmz_file_paths = []
with open(input_file_path, 'r') as file:
    kmz_file_paths = [line.strip() for line in file]

# Process each KMZ file path
updated_kmz_file_paths = []
for kmz_path in kmz_file_paths:
    updated_path = process_kmz_file(kmz_path)
    if updated_path:
        updated_kmz_file_paths.append(updated_path)

# Save the updated file paths to a new .txt document
output_file_path = os.path.join(os.path.expanduser('~'), 'Downloads', 'updated_kmz_paths.txt')
with open(output_file_path, 'w') as output_file:
    output_file.write('\n'.join(updated_kmz_file_paths))

print(f"Processing complete. Updated KMZ file paths saved to '{output_file_path}'.")
