import os

# Directory path
directory_path = r'C:\Users\mmccarthy\Desktop\Automation Team\Lightsett\MRSummary Pull\MakeReadySummaryAttributes'

# List all files in the directory
files = os.listdir(directory_path)

# Iterate through each file
for file_name in files:
    if "MakeReady" in file_name:
        # Split the file name by "MakeReady" and get the extension
        base_name, extension = os.path.splitext(file_name)
        new_file_name = base_name.split("MakeReady")[0] + "MakeReady" + extension
        
        # Get the full path of the file
        file_path = os.path.join(directory_path, file_name)
        new_file_path = os.path.join(directory_path, new_file_name)
        
        try:
            # Rename the file
            os.rename(file_path, new_file_path)
            print(f"Renamed {file_name} to {new_file_name}")
        except Exception as e:
            print(f"Failed to rename {file_name}: {e}")
    else:
        print(f"File {file_name} does not contain 'MakeReady'")

print("Renaming complete.")
