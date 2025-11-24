import os

directory_to_clear = '/content/'

# Get a list of all files and directories in /content
items_in_content = os.listdir(directory_to_clear)

print(f"Files and directories in {directory_to_clear} before clearing:")
for item in items_in_content:
    print(item)

for item_name in items_in_content:
    item_path = os.path.join(directory_to_clear, item_name)
    try:
        if os.path.isfile(item_path): # Check if it's a file before removing
            os.remove(item_path)
            print(f"Removed file: {item_path}")
        elif os.path.isdir(item_path): # Handle directories, if necessary (e.g., delete recursively)
            # For now, we'll just skip directories to prevent accidental deletion of important Colab folders like sample_data
            print(f"Skipping directory: {item_path}")
    except Exception as e:
        print(f"Error removing {item_path}: {e}")

print(f"\nFiles and directories in {directory_to_clear} after clearing:")
# List remaining items to confirm
remaining_items = os.listdir(directory_to_clear)
if remaining_items:
    for item in remaining_items:
        print(item)
else:
    print("Directory is empty (except for any skipped directories).")
