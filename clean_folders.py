import os
import tkinter as tk
from tkinter import filedialog

def delete_empty_folders():
    # 1. Setup UI
    root = tk.Tk()
    root.withdraw()

    print("--- 📂 Empty Folder Cleanup Tool ---")
    
    # 2. Select the Main Folder
    target_dir = filedialog.askdirectory(title="Select the Folder to Clean (Removes all empty subfolders)")
    
    if not target_dir:
        print("Selection cancelled.")
        return

    deleted_count = 0

    # 3. Walk through folders from bottom to top (topdown=False)
    # This ensures we delete empty sub-folders before checking if the parent is empty.
    for root_path, dirs, files in os.walk(target_dir, topdown=False):
        for name in dirs:
            full_path = os.path.join(root_path, name)
            
            # Check if the folder is empty
            try:
                if not os.listdir(full_path):
                    os.rmdir(full_path)
                    print(f"🗑️ Deleted empty folder: {full_path}")
                    deleted_count += 1
            except Exception as e:
                print(f"⚠️ Could not delete {name}: {e}")

    print(f"\n✅ Cleanup Complete! Total empty folders removed: {deleted_count}")
    input("Press Enter to close...")

if __name__ == "__main__":
    delete_empty_folders()