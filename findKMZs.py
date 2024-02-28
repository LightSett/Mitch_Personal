import os

# Function to find all KMZ files within a directory and its subdirectories
def find_kmz_files(directory):
    kmz_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.kmz'):
                kmz_path = os.path.join(root, file)
                kmz_files.append(kmz_path)
    return kmz_files

# Specify the directory to search for KMZ files
search_directory = r'S:\Customers\48 Glass Roots\48.001.001 Evans, CO\Evans, CO\2. LLD'

# Call the function to find KMZ files
kmz_files = find_kmz_files(search_directory)

# Determine the user's Downloads folder
downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')

# Create the output file path in the Downloads folder
output_file_path = os.path.join(downloads_folder, 'kmz_file_list_new.txt')

# Write the list of KMZ file paths to the output file
with open(output_file_path, 'w') as output_file:
    for kmz_file in kmz_files:
        output_file.write(kmz_file + '\n')

print(f"List of KMZ files saved to: {output_file_path}")
