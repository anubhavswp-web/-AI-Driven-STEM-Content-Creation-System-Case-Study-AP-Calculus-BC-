# create_drive_folders_local.py
# Fixed: also reads subtopic in the same row as the main topic (e.g., A2 & B2)
# Run: python "D:\New project\Sub folder creater\create_drive_folders_local.py"

import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import unicodedata

# -------------------------
# sanitize and cleanup
# -------------------------
def replace_and_strip(s: str) -> str:
    if s is None:
        return ''
    s = str(s).strip()
    # replace colon as requested
    s = s.replace(':', ' -')
    # replace slashes and pipe etc with hyphen or apostrophe
    repl_map = {
        '/': '-',
        '\\': '-',
        '|': '-',
        '"': "'",
    }
    for k, v in repl_map.items():
        s = s.replace(k, v)
    # remove characters with Unicode category not in L,N,P,Z (letters, numbers, punctuation, separators)
    filtered_chars = []
    for ch in s:
        cat = unicodedata.category(ch)
        if cat and cat[0] in ('L', 'N', 'P', 'Z'):
            filtered_chars.append(ch)
        # else drop the character
    s = ''.join(filtered_chars)
    # Remove any remaining illegal characters just in case
    s = re.sub(r'[<>:"/\\|?*]', '', s)
    # collapse multiple spaces
    s = re.sub(r'\s+', ' ', s).strip()
    # trim trailing dots/spaces (Windows limitation)
    s = s.rstrip(' .')
    if s == '':
        return 'Unnamed'
    # Avoid Windows reserved names
    base_upper = os.path.splitext(s)[0].upper()
    RESERVED = {"CON","PRN","AUX","NUL"} | {f"COM{i}" for i in range(1,10)} | {f"LPT{i}" for i in range(1,10)}
    if base_upper in RESERVED:
        s = s + '_folder'
    return s

# -------------------------
# read first two columns (fixed)
# -------------------------
def read_first_two(path):
    p = str(path)
    low = p.lower()
    # read without treating first row as header
    if low.endswith('.csv'):
        df = pd.read_csv(p, header=None, dtype=str, keep_default_na=False)
    elif low.endswith(('.xls', '.xlsx')):
        df = pd.read_excel(p, header=None, dtype=str)
    else:
        df = pd.read_csv(p, header=None, dtype=str, keep_default_na=False)
    # Only first two columns
    if df.shape[1] == 0:
        return pd.DataFrame(columns=[0,1])
    df = df.iloc[:, :2]
    # fill missing and strip whitespace
    df = df.fillna('').astype(str)
    df = df.apply(lambda col: col.str.strip())
    return df

# -------------------------
# safe folder creation
# -------------------------
def ensure_folder(path):
    if not os.path.exists(path):
        os.makedirs(path)
        return True
    return False

# -------------------------
# main routine
# -------------------------
def main():
    print("\n=== Folder Creator (local) ===\n")
    # pick file
    root_tk = tk.Tk()
    root_tk.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select one CSV / XLSX file",
        filetypes=[("Excel & CSV files", "*.xlsx *.xls *.csv")]
    )
    if not file_path:
        print("No file selected. Exiting.")
        return

    # read file (no header) and then drop only the sheet header (Excel row 1)
    df = read_first_two(file_path)
    if df.empty:
        print("File appears to have no data (after reading first two columns). Exiting.")
        return

    # As you specified: Row 1 is always a header -> drop it (this leaves Excel row2 as df_data row0)
    df_data = df.iloc[1:].reset_index(drop=True)
    if df_data.empty:
        print("No data rows after dropping header. Exiting.")
        return

    # ask for root folder (default provided)
    default_root = r"G:\My Drive\AP Calculus\AP Calculus BC final\AP calculus Study guides"
    print(f"\nSelected file: {file_path}")
    root_input = input(f"\nEnter LOCAL ROOT folder path (press Enter to use default):\n[{default_root}]\n> ").strip().strip('"')
    if root_input == '':
        root_folder = default_root
    else:
        root_folder = root_input
    root_folder = os.path.normpath(root_folder)

    # parent folder name = file stem (without extension)
    file_stem = Path(file_path).stem
    parent_raw = file_stem
    parent_safe = replace_and_strip(parent_raw)
    parent_path = os.path.join(root_folder, parent_safe)

    # create parent
    ensure_folder(root_folder)
    created_parent = ensure_folder(parent_path)
    print(f"\nParent folder: {parent_path}   (created={created_parent})")

    # Prepare to iterate rows: df_data rows correspond to Excel rows starting from row2
    n = len(df_data)
    idx = 0
    main_count = 0
    sub_count = 0

    while idx < n:
        raw_main = str(df_data.iat[idx, 0]).strip()
        # if main empty, skip forward until next main or end (we must find main to start)
        if raw_main == '' or raw_main.lower() == 'nan':
            idx += 1
            continue

        # found a main topic
        main_count += 1
        main_raw = raw_main
        main_safe = replace_and_strip(main_raw)
        main_path = os.path.join(parent_path, main_safe)
        created = ensure_folder(main_path)
        print(f"{'Created' if created else 'Exists '} MAIN -> {main_safe}    (original: {main_raw})")

        # *** NEW: check subtopic in the SAME ROW as the main (e.g., A2 & B2)
        raw_sub_same_row = str(df_data.iat[idx, 1]).strip()
        if raw_sub_same_row != '' and raw_sub_same_row.lower() != 'nan':
            sub_safe0 = replace_and_strip(raw_sub_same_row)
            sub_path0 = os.path.join(main_path, sub_safe0)
            created_sub0 = ensure_folder(sub_path0)
            sub_count += 1
            print(f"    {'Created' if created_sub0 else 'Exists '} SUB  -> {sub_safe0}    (original: {raw_sub_same_row})")

        # collect subtopics scanning downward up to 50 rows lookahead
        blanks = 0
        lookahead_limit = 50
        look_idx = idx + 1
        steps = 0
        while look_idx < n and steps < lookahead_limit:
            # if next row has a new main topic (non-empty A), stop
            next_A = str(df_data.iat[look_idx, 0]).strip()
            if next_A != '' and next_A.lower() != 'nan':
                # reachable next main detected: break to outer loop (we will set idx to this row)
                break

            # check column B for subtopic
            raw_sub = str(df_data.iat[look_idx, 1]).strip()
            if raw_sub != '' and raw_sub.lower() != 'nan':
                # valid subtopic - create
                sub_safe = replace_and_strip(raw_sub)
                sub_path = os.path.join(main_path, sub_safe)
                created_sub = ensure_folder(sub_path)
                sub_count += 1
                print(f"    {'Created' if created_sub else 'Exists '} SUB  -> {sub_safe}    (original: {raw_sub})")
                blanks = 0  # reset blanks counter
            else:
                # empty B cell
                blanks += 1
                if blanks > 5:
                    # stop collecting subtopics for this main
                    break

            look_idx += 1
            steps += 1

        # advance idx to the next main if found within lookahead; otherwise set idx = look_idx
        next_main_idx = None
        for j in range(idx + 1, min(n, idx + 1 + lookahead_limit)):
            if str(df_data.iat[j, 0]).strip() != '' and str(df_data.iat[j, 0]).strip().lower() != 'nan':
                next_main_idx = j
                break

        if next_main_idx is not None:
            idx = next_main_idx
        else:
            idx = look_idx

    print(f"\nDone. Created/verified {main_count} main folders and {sub_count} subfolders (approx).")
    print(f"Parent folder is: {parent_path}")
    print("\nScript finished.\n")

if __name__ == "__main__":
    main()
