import os
import shutil

# Function to check if a file has no date at the end of its name
def has_no_date(filename):
    # Modify this condition as needed
    return not filename[-6:-4].isdigit()

# Function to move .KMZ files from _Archive folders to the KMZ folder
def move_kmz_files(directory):
    kmz_folder = os.path.join(directory, "KMZ")

    # Create the KMZ folder if it doesn't exist
    if not os.path.exists(kmz_folder):
        os.makedirs(kmz_folder)

    for root, dirs, files in os.walk(directory):
        for dir_name in dirs:
            if dir_name == "_Archive":
                archive_folder = os.path.join(root, dir_name)
                for filename in os.listdir(archive_folder):
                    if filename.endswith(".KMZ") and has_no_date(filename):
                        source_path = os.path.join(archive_folder, filename)
                        destination_path = os.path.join(kmz_folder, filename)

                        # Move the .KMZ file to the KMZ folder
                        shutil.move(source_path, destination_path)
                        print(f"Moved {filename} to KMZ folder.")

if __name__ == "__main__":
    target_directory = r"S:\Customers\48 Glass Roots\48.001.001 Evans, CO\Evans, CO\2. LLD"

    if os.path.exists(target_directory):
        move_kmz_files(target_directory)
        print("Finished moving KMZ files.")
    else:
        print("Directory does not exist.")
