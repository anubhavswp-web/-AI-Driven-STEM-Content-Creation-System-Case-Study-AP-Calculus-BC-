import os
import tkinter as tk
from tkinter import filedialog, messagebox

def delete_files_by_prefix():
    # 1. Setup a hidden root window for the folder picker
    root = tk.Tk()
    root.withdraw()

    # 2. Ask user to select the directory
    target_folder = filedialog.askdirectory(title="Select the Root Folder")
    if not target_folder:
        print("No folder selected. Exiting.")
        return

    # 3. Ask user for the prefix (the start of the filename)
    prefix = input(f"Enter the starting name (prefix) of files to delete in '{target_folder}': ").strip()
    
    if not prefix:
        print("Prefix cannot be empty.")
        return

    confirm = input(f"Are you sure you want to delete all files starting with '{prefix}'? (yes/no): ")
    if confirm.lower() != 'yes':
        print("Operation cancelled.")
        return

    deleted_count = 0

    # 4. Walk through the folder and nested subfolders
    for root_dir, dirs, files in os.walk(target_folder):
        for filename in files:
            if filename.startswith(prefix):
                file_path = os.path.join(root_dir, filename)
                try:
                    os.remove(file_path)
                    print(f"Deleted: {file_path}")
                    deleted_count += 1
                except Exception as e:
                    print(f"Error deleting {file_path}: {e}")

    print(f"\nTask Complete. Total files deleted: {deleted_count}")
    messagebox.showinfo("Done", f"Deleted {deleted_count} files.")

if __name__ == "__main__":
    delete_files_by_prefix()