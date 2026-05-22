import os
import pandas as pd  # FIXED THE TYPO HERE
from google import genai
import subprocess
import tkinter as tk
from tkinter import filedialog
import time
import re
import threading
import keyboard

# --- CONFIGURATION ---
PROMPT_DIR = r"D:\New project\To full auto\Prompts\AutoFullSwing"
TEMPLATE_PATH = os.path.join(PROMPT_DIR, "template.docx")
API_KEY_PATH = os.path.join(PROMPT_DIR, "api_key.txt")

is_paused = False
stop_program = False
MIN_SIZE_BYTES = 10 * 1024 
files_skipped = 0
files_generated = 0

def clean_path(text):
    """Removes invalid Windows characters and trailing dots/spaces."""
    text = str(text).strip()
    text = re.sub(r'[<>:"/\\|?*]', '', text)
    return text.strip('. ')

def get_prompt(filename):
    path = os.path.join(PROMPT_DIR, filename)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return ""

def generate_doc(client, final_prompt, folder, filename):
    global files_skipped, files_generated
    file_path = os.path.join(folder, f"{filename}.docx")
    
    # SMART SKIP LOGIC
    if os.path.exists(file_path) and os.path.getsize(file_path) >= MIN_SIZE_BYTES:
        files_skipped += 1
        print(f"   ⏭️ Skipping: {filename}")
        return True

    temp_md = f"temp_{int(time.time())}.md"
    
    for attempt in range(3):
        if stop_program: return False
        while is_paused: time.sleep(0.5)
        
        try:
            time.sleep(12) # Tier 1 Pacing
            print(f"   ⚙️ Generating: {filename}...")
            
            response = client.models.generate_content(
                model="gemini-2.0-flash", 
                contents=final_prompt
            )
            
            text = response.text.replace("```markdown", "").replace("```", "").strip()
            with open(temp_md, "w", encoding="utf-8") as f:
                f.write(text)
                
            subprocess.run(["pandoc", temp_md, "--reference-doc", TEMPLATE_PATH, "-o", file_path], capture_output=True)
            
            if os.path.exists(file_path):
                files_generated += 1
                if os.path.exists(temp_md): os.remove(temp_md)
                return True
        except Exception as e:
            print(f"   ⚠️ Error: {str(e)[:50]}. Retrying...")
            time.sleep(15)
    return False

def run_main():
    global stop_program, files_skipped, files_generated
    
    with open(API_KEY_PATH, "r") as f: key = f.read().strip()
    client = genai.Client(api_key=key)

    root = tk.Tk(); root.withdraw()
    excel_file = filedialog.askopenfilename(title="Select your Excel file")
    save_destination = filedialog.askdirectory(title="Select where to create the Unit Folder")
    
    if not excel_file or not save_destination: return

    # 1. CREATE MAIN UNIT FOLDER (Named after Excel)
    excel_name = os.path.splitext(os.path.basename(excel_file))[0]
    unit_root = os.path.join(save_destination, clean_path(excel_name))
    os.makedirs(unit_root, exist_ok=True)

    # 2. LOAD PROMPTS
    master = get_prompt("1. Master prompt.txt")
    p_unit_guide = get_prompt("Study_Guide_Prompt.txt") # Theoretical Unit Guide
    p_topic_guide = get_prompt("Subguide_Prompt.txt")   # Main Topic Guide
    p_notes = get_prompt("Notes_Prompt.txt")
    p_ex = get_prompt("4. Prompt for Examples.txt")
    p_mcq = get_prompt("5. Prompt for MCQs.txt")
    p_frq = get_prompt("6. Prompts for FRQs.txt")
    p_prac = get_prompt("7. Prompt for Practice questions.txt")
    p_test = get_prompt("8. Prompt for Practice test.txt")

    threading.Thread(target=keyboard_listener, daemon=True).start()

    df = pd.read_excel(excel_file)
    df.iloc[:, 0:2] = df.iloc[:, 0:2].ffill()
    
    print(f"\n🚀 Creating Unit: {excel_name}")

    # --- TIER 1: UNIT STUDY GUIDE (Inside Main Folder) ---
    unit_context = df.to_string()
    unit_prompt = f"{master}\n\n{p_unit_guide}\n\nUNIT DATA:\n{unit_context}"
    generate_doc(client, unit_prompt, unit_root, f"00_Complete_Unit_Study_Guide")

    # --- TIER 2: MAIN TOPICS ---
    for topic_id, t_group in df.groupby(df.columns[0], sort=False):
        if stop_program: break
        
        t_folder = os.path.join(unit_root, clean_path(topic_id))
        os.makedirs(t_folder, exist_ok=True)
        
        print(f"\n📂 Main Topic: {topic_id}")
        
        # Generate Main Topic Guide
        topic_context = t_group.to_string()
        topic_prompt = f"{master}\n\n{p_topic_guide}\n\nTOPIC DATA:\n{topic_context}"
        generate_doc(client, topic_prompt, t_folder, f"01_Topic_Study_Guide")

        # --- TIER 3: SUBTOPICS ---
        for sub_id, s_group in t_group.groupby(df.columns[1], sort=False):
            if stop_program: break
            
            s_folder = os.path.join(t_folder, clean_path(sub_id))
            os.makedirs(s_folder, exist_ok=True)
            
            s_details = "\n".join(s_group.iloc[:, 2].astype(str).tolist())

            tasks = [
                ("Notes", p_notes), ("Examples", p_ex), ("MCQs", p_mcq),
                ("FRQs", p_frq), ("Practice", p_prac), ("Test", p_test)
            ]

            for prefix, task_p in tasks:
                if stop_program: break
                final_p = f"{master}\n\n{task_p.replace('{{SUBTOPIC}}', str(sub_id)).replace('{{DETAILS}}', s_details)}"
                generate_doc(client, final_p, s_folder, f"{prefix}_{clean_path(sub_id)}")

    print(f"\n✅ FINISHED. Created folders and guides for {excel_name}")
    input("Press Enter to close...")

def keyboard_listener():
    global is_paused, stop_program
    while not stop_program:
        if keyboard.is_pressed('p'): is_paused = True
        if keyboard.is_pressed('r'): is_paused = False
        if keyboard.is_pressed('esc'): stop_program = True
        time.sleep(0.1)

if __name__ == "__main__":
    run_main()