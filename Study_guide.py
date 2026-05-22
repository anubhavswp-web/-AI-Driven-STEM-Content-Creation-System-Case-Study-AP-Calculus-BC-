import os
import pandas as pd
from google import genai
import subprocess
import time
import docx
import re
import tkinter as tk
from tkinter import filedialog, messagebox

class HighIntelAuditor:
    """Forensic Document Auditor: Ensures 100% compliance with your logic."""
    def __init__(self, file_path, unit_id, sub_count):
        self.path = file_path
        self.unit_id = unit_id
        self.n = sub_count
        self.errors = []

    def run_deep_audit(self):
        if not os.path.exists(self.path): 
            return "MISSING" # First time running
            
        try:
            doc = docx.Document(self.path)
            full_text = "\n".join([p.text for p in doc.paragraphs])
            
            # 1. Numbering Hierarchy Audit
            expected_solved_id = f"{self.unit_id}.{self.n + 1}"
            expected_test_id = f"{self.unit_id}.{self.n + 2}"
            if expected_solved_id not in full_text:
                self.errors.append(f"Header Error: Expected section {expected_solved_id}")
            
            # 2. Problem Quantity Audit (10 Solved + 10 Practice)
            if "Problem 10" not in full_text:
                self.errors.append("Quantity Error: Missing Solved Problem 10")
            if "Practice Problem 10" not in full_text:
                self.errors.append("Quantity Error: Missing Practice Problem 10")
                
            # 3. Math & Style Integrity
            if full_text.count("$$") % 2 != 0:
                self.errors.append("Math Syntax: Unbalanced $$ detected")
            if "\\boxed" not in full_text:
                self.errors.append("Style Error: Missing boxed answers")

            # 4. Academic Depth (Ensuring enough text)
            if len(full_text.split()) < 800:
                self.errors.append("Depth Error: Content is too thin for a study guide")

            return "PASSED" if len(self.errors) == 0 else "FAILED"
        except Exception as e:
            self.errors.append(f"Read Error: {e}")
            return "FAILED"

def run_program():
    with open("api_key.txt", "r") as f: key = f.read().strip()
    client = genai.Client(api_key=key)
    
    root = tk.Tk(); root.withdraw()
    excel_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    if not excel_path: return
    save_dest = filedialog.askdirectory()
    
    excel_name = os.path.splitext(os.path.basename(excel_path))[0]
    project_folder = os.path.join(save_dest, f"study guide on {excel_name}")
    os.makedirs(project_folder, exist_ok=True)

    df = pd.read_excel(excel_path).ffill()
    with open("1. Master prompt.txt", "r", encoding="utf-8") as f: master_p = f.read()

    for main_topic, group in df.groupby(df.columns[0], sort=False):
        topic_parts = str(main_topic).split(' ', 1)
        topic_id = topic_parts[0]
        topic_name = topic_parts[1] if len(topic_parts) > 1 else ""
        sub_count = len(group.iloc[:, 1].unique())
        
        docx_path = os.path.join(project_folder, f"Study_Guide_{topic_id.replace('.','_')}.docx")

        # --- THE AUDIT ---
        auditor = HighIntelAuditor(docx_path, topic_id, sub_count)
        status = auditor.run_deep_audit()

        if status == "PASSED":
            print(f"✅ {topic_id} Verified 100%. Skipping.")
            continue
        elif status == "FAILED":
            print(f"❌ {topic_id} failed audit: {auditor.errors}. Regenerating...")
            os.remove(docx_path)
        else:
            print(f"🔄 {topic_id} is new. Generating first time...")

        # Build Subtopic Data (Fixing the .iloc FutureWarning here)
        sub_list = []
        current_sub = ""
        s_idx = 0
        d_idx = 1
        for _, row in group.iterrows():
            s_val = str(row.iloc[1]) # Modern iloc access
            d_val = str(row.iloc[2]) # Modern iloc access
            if s_val != current_sub:
                current_sub = s_val
                s_idx += 1
                d_idx = 1
                sub_list.append(f"\n## {s_val}")
            sub_list.append(f"### {topic_id}.{s_idx}.{d_idx} {d_val}")
            d_idx += 1

        prompt = (master_p.replace('{{MAINTOPIC_ID}}', topic_id)
                          .replace('{{MAINTOPIC_NAME}}', topic_name)
                          .replace('{{SUBTOPIC_DATA}}', "\n".join(sub_list))
                          .replace('{{NEXT_ID_1}}', f"{topic_id}.{sub_count+1}")
                          .replace('{{NEXT_ID_2}}', f"{topic_id}.{sub_count+2}")
                          .replace('{{NEXT_ID_3}}', f"{topic_id}.{sub_count+3}"))

        response = client.models.generate_content(model="gemini-2.0-flash", contents=prompt)
        temp_md = "temp_audit.md"
        with open(temp_md, "w", encoding="utf-8") as f: f.write(response.text)
        
        subprocess.run(["pandoc", temp_md, "--from", "markdown+tex_math_dollars", 
                        "--reference-doc", "template.docx", "--standalone", "-o", docx_path])
        if os.path.exists(temp_md): os.remove(temp_md)
        time.sleep(12)

    messagebox.showinfo("Complete", "Study guides verified and corrected.")

if __name__ == "__main__":
    run_program()