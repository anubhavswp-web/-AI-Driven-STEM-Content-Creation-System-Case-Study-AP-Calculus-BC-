# ===================== SAFE BLANK FILE CLEANER =====================
# GUARANTEED TO RUN VERSION
# No external libraries required
# ==================================================================

import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import traceback


def main():
    try:
        print("Script started")

        # ---- Tkinter setup ----
        root = tk.Tk()
        root.withdraw()

        base_folder = filedialog.askdirectory(
            title="Select ROOT folder to scan for blank files"
        )

        if not base_folder:
            print("No folder selected. Exiting.")
            input("Press Enter to exit...")
            return

        print("Selected folder:", base_folder)

        BLANK_FOLDER_NAME = "_BLANK_FILES"
        MAX_DOCX_SIZE_KB = 11

        blank_folder = os.path.join(base_folder, BLANK_FOLDER_NAME)
        os.makedirs(blank_folder, exist_ok=True)

        total_files = 0
        blank_files = 0

        # ---- Scan folders ----
        for root_dir, _, files in os.walk(base_folder):
            if BLANK_FOLDER_NAME in root_dir:
                continue

            for filename in files:
                file_path = os.path.join(root_dir, filename)
                ext = filename.lower()

                # -------- Word (.docx) --------
                if ext.endswith(".docx"):
                    total_files += 1
                    size_kb = os.path.getsize(file_path) / 1024

                    if size_kb <= MAX_DOCX_SIZE_KB:
                        move_file(file_path, blank_folder)
                        blank_files += 1

                # -------- TXT / MD --------
                elif ext.endswith((".txt", ".md")):
                    total_files += 1
                    try:
                        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                            content = f.read()
                        if not content.strip():
                            move_file(file_path, blank_folder)
                            blank_files += 1
                    except Exception:
                        pass

        summary = (
            f"SCAN COMPLETE\n"
            f"------------------\n"
            f"Total files scanned : {total_files}\n"
            f"Blank files moved   : {blank_files}\n\n"
            f"Moved to:\n{blank_folder}"
        )

        print(summary)
        messagebox.showinfo("Done", summary)

        input("Press Enter to exit...")

    except Exception as e:
        print("\n❌ ERROR OCCURRED")
        traceback.print_exc()
        input("Press Enter to exit...")


def move_file(src, dest_folder):
    filename = os.path.basename(src)
    dest = os.path.join(dest_folder, filename)

    base, ext = os.path.splitext(filename)
    counter = 1
    while os.path.exists(dest):
        dest = os.path.join(dest_folder, f"{base}_{counter}{ext}")
        counter += 1

    shutil.move(src, dest)


if __name__ == "__main__":
    main()
