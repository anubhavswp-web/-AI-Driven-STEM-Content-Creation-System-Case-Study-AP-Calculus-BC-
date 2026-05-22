import os
import re
import json
import textwrap
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from google import genai
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime

# --- MASTER CONFIGURATION ---

# ============================================================
#  OMML (Word Math Equation) -> Plain Text Extractor
#  ROOT CAUSE FIX: p.text silently drops all <m:oMath> math nodes,
#  so the AI received blank equations and invented random functions.
# ============================================================
MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
W_NS    = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def _omml_node_to_text(elem):
    local = elem.tag.split("}")[1] if "}" in elem.tag else elem.tag
    def ct(e):
        if e is None:
            return ""
        return "".join(_omml_node_to_text(c) for c in e)
    if local in ("oMath", "oMathPara"):
        return ct(elem)
    elif local == "r":
        return "".join(t.text or "" for t in elem.findall("{" + MATH_NS + "}t"))
    elif local == "sSup":
        b = elem.find("{" + MATH_NS + "}e")
        s = elem.find("{" + MATH_NS + "}sup")
        return ct(b) + "^" + ct(s)
    elif local == "sSub":
        b = elem.find("{" + MATH_NS + "}e")
        s = elem.find("{" + MATH_NS + "}sub")
        return ct(b) + "_" + ct(s)
    elif local == "sSubSup":
        b  = elem.find("{" + MATH_NS + "}e")
        sb = elem.find("{" + MATH_NS + "}sub")
        sp = elem.find("{" + MATH_NS + "}sup")
        return ct(b) + "_" + ct(sb) + "^" + ct(sp)
    elif local == "f":
        n = elem.find("{" + MATH_NS + "}num")
        d = elem.find("{" + MATH_NS + "}den")
        return "(" + ct(n) + ")/(" + ct(d) + ")"
    elif local == "rad":
        e_node   = elem.find("{" + MATH_NS + "}e")
        deg_node = elem.find("{" + MATH_NS + "}deg")
        inner = ct(e_node)
        deg   = ct(deg_node)
        return ("(" + inner + ")^(1/" + deg + ")") if (deg_node is not None and deg.strip()) else ("sqrt(" + inner + ")")
    elif local == "d":
        parts = [ct(e) for e in elem.findall("{" + MATH_NS + "}e")]
        return "(" + ",".join(parts) + ")"
    elif local == "func":
        fn  = elem.find("{" + MATH_NS + "}fName")
        arg = elem.find("{" + MATH_NS + "}e")
        return ct(fn) + "(" + ct(arg) + ")"
    elif local == "nary":
        chr_e = elem.find(".//{" + MATH_NS + "}chr")
        sym   = chr_e.get("{" + MATH_NS + "}val", "integral") if chr_e is not None else "integral"
        sub_e = elem.find("{" + MATH_NS + "}sub")
        sup_e = elem.find("{" + MATH_NS + "}sup")
        e_e   = elem.find("{" + MATH_NS + "}e")
        return sym + "_" + ct(sub_e) + "^" + ct(sup_e) + " " + ct(e_e)
    elif local == "limLow":
        e_e   = elem.find("{" + MATH_NS + "}e")
        lim_e = elem.find("{" + MATH_NS + "}lim")
        return ct(e_e) + "_" + ct(lim_e)
    else:
        return "".join(_omml_node_to_text(c) for c in elem)


def extract_paragraph_text(para):
    """Full paragraph text including OMML math equations (which p.text drops)."""
    parts = []
    for child in para._p:
        local = child.tag.split("}")[1] if "}" in child.tag else child.tag
        if local == "r":
            parts.append("".join(t.text or "" for t in child.findall("{" + W_NS + "}t")))
        elif local in ("oMath", "oMathPara"):
            parts.append(_omml_node_to_text(child))
        elif local in ("hyperlink", "ins", "del"):
            for r in child.findall(".//{" + W_NS + "}r"):
                parts.append("".join(t.text or "" for t in r.findall("{" + W_NS + "}t")))
    return "".join(parts)
# ============================================================

BG_COLOR = "#F0F8FF"    # Very Light Blue
DARK_BLUE = "#00008B"   # Sharp Dark Blue
RED_ACCENT = "#FF0000"  # For Holes/Asymptotes
CAPTION_RED = "#C41E3A"  # Deep Crimson Red for description text

def load_api_key():
    try:
        with open("api_key.txt", "r") as f:
            return f.read().strip()
    except FileNotFoundError:
        return None

client = genai.Client(api_key=load_api_key())

# --- SYSTEM INSTRUCTION ---
SYSTEM_INSTRUCTION = r"""
Role: Harvard Math Professor & Technical Illustrator.
Mission: Parse AP Calculus docs and identify MORE than 10 locations for visuals.

STRICT VISUAL RULES:
1. BG: #F0F8FF. TEXT/AXES: #00008B.
2. Coordinate axes with division tick marks on BOTH axes.
3. STROKE: Primary functions = 3.0pt.
4. INDICATORS: Red (#FF0000) for Holes and Asymptotes (dashed).
5. LABELS: Equations close to curves in Dark Blue.
6. CRITICAL POINTS: Mark ALL intersection points with coordinates.
7. ASYMPTOTES: Red dashed lines with equations labeled.
8. Single color per function, different colors for different functions.
9. Sharp image - readable at 50% scale.

CODE RULES:
- ONLY use: plt.plot(), plt.scatter(), plt.axhline(), plt.axvline(), plt.text()
- NO: canvas, fig.canvas, FigureCanvas, draw_idle, blit, add_subplot
- For asymptotes: plt.axvline(x, linestyle='--', color='#FF0000', linewidth=2.0)
- For holes: plt.scatter(x, y, s=150, facecolors='none', edgecolors='#FF0000', linewidths=3.0)
- Always: plt.grid(True, alpha=0.3, color='#00008B')
- Label axes: plt.xlabel('x', color='#00008B', fontsize=13, weight='bold')

DESCRIPTION RULES - CRITICAL:

YOU MUST WRITE 3-5 LINE DESCRIPTIONS for each image.this is very important and the strict rule .

This text will appear BELOW the graph (no box, just text) to help students understand.

FORMAT: 3-5 lines separated by \n
Line 1: Section/Problem/Theorem reference
Lines 2-4: Explain mathematical concept
Line 5: What to observe

NO AI LANGUAGE: Never "This graph shows", "The figure illustrates"
MUST REFERENCE: Specific section from document
MUST EXPLAIN: Mathematical concept in detail

EXAMPLES:

✅ PERFECT:
"Section 1.1: Graphical Interpretation of Limits
The limit describes y-value function approaches as x approaches c.
Independent of whether f(c) is defined at that point.
Observe convergence from both left and right directions.
Demonstrates two-sided limit behavior graphically."

✅ PERFECT:
"Problem 3.4: Vertical Asymptotes in Rational Functions
Asymptotes occur where denominator equals zero, numerator non-zero.
Function approaches ±∞ at these x-values shown by red dashed lines.
Notice unbounded behavior at x=2 and x=-3 with opposite directions."

OUTPUT JSON FORMAT with complete example:
[
  {
    "marker": "limit_1",
    "code": "x = np.linspace(-3, 3, 500)\ny = x**2\nplt.plot(x, y, color='#00008B', linewidth=3.0)\nplt.axhline(0, color='#00008B', linewidth=1)\nplt.axvline(0, color='#00008B', linewidth=1)\nplt.grid(True, alpha=0.3, color='#00008B')\nplt.xlabel('x', color='#00008B', fontsize=13, weight='bold')\nplt.ylabel('y', color='#00008B', fontsize=13, weight='bold')",
    "description": "Section 1.2: Quadratic Function Graph\nParabola y=x² with vertex at origin (0,0).\nSymmetric about y-axis with minimum value at vertex.\nFunction increases without bound as |x| increases.\nDemonstrates basic shape of quadratic functions.",
    "para_index": 10
  }
]

The "description" field MUST have 3-5 lines of educational text.
"""


def sanitize_json_str(s):
    """Remove invalid JSON escape sequences that Gemini sometimes produces."""
    # Replace backslash followed by any char that isn't a valid JSON escape
    result = []
    i = 0
    while i < len(s):
        if s[i] == '\\' and i + 1 < len(s):
            next_c = s[i + 1]
            if next_c in ('"', '\\', '/', 'b', 'f', 'n', 'r', 't', 'u'):
                result.append(s[i])
                result.append(next_c)
                i += 2
            else:
                # Invalid escape — double the backslash to make it valid
                result.append('\\\\')
                i += 1
        else:
            result.append(s[i])
            i += 1
    return ''.join(result)

class STEMOrchestrator:
    def process_file(self, doc_path):
        doc = Document(doc_path)
        content_lines = []
        for i, p in enumerate(doc.paragraphs):
            full_text = extract_paragraph_text(p)
            if full_text.strip():
                content_lines.append("[" + str(i) + "] " + full_text)
        content = "\n".join(content_lines)
        
        try:
            response = client.models.generate_content(
                model='gemini-3-flash-preview',
                contents=f"""Parse this AP Calculus document and create visualizations with descriptions.

CRITICAL: Write 3-5 line descriptions for EVERY image. This text appears below the graph to help students.

Example description format:
"Section 1.1: Limit Introduction\nDescribes value function approaches as x nears a point.\nIndependent of actual function value at that point.\nObserve approach from both directions.\nDemonstrates fundamental limit concept."

Document:
{content}""",
                config={'system_instruction': SYSTEM_INSTRUCTION}
            )
            
            clean_json = re.search(r'\[.*\]', response.text, re.DOTALL)
            if not clean_json:
                print("ERROR: No JSON found")
                return "AI did not return valid JSON"
                
            tasks = json.loads(sanitize_json_str(clean_json.group()))
            print(f"✓ Generated {len(tasks)} image tasks")
            
            # Validate descriptions
            for i, task in enumerate(tasks):
                desc = task.get('description', '').strip()
                if len(desc) < 50:
                    print(f"⚠ WARNING: Task {i} has SHORT description ({len(desc)} chars)")
                else:
                    lines = desc.count('\n') + 1
                    print(f"✓ Task {i}: {len(desc)} chars, {lines} lines")
                    
        except Exception as e:
            print(f"ERROR: {e}")
            return f"AI Error: {str(e)}"

        output_dir = os.path.join(os.path.dirname(doc_path), "Completed_Units")
        os.makedirs(output_dir, exist_ok=True)

        successful = 0
        failed = 0
        no_description = 0
        
        for task in sorted(tasks, key=lambda x: x['para_index'], reverse=True):
            img_name = f"temp_{task['para_index']}.png"
            
            # Validate code
            forbidden_terms = ['canvas', 'FigureCanvas', 'draw_idle', 'blit', 'add_subplot']
            if any(term.lower() in task['code'].lower() for term in forbidden_terms):
                print(f"⚠ Skipped para {task['para_index']}: forbidden term")
                failed += 1
                continue
            
            # Get description
            description_text = task.get('description', '').strip()
            has_description = len(description_text) >= 50
            
            if not has_description:
                print(f"⚠ Para {task['para_index']}: No description - creating without text")
                no_description += 1
            
            try:
                # Wrap description and count lines
                if has_description:
                    wrapped_lines = textwrap.wrap(description_text, width=70)
                    if len(wrapped_lines) > 5:
                        wrapped_lines = wrapped_lines[:5]  # Max 5 lines
                    
                    num_lines = len(wrapped_lines)
                    wrapped_text = "\n".join(wrapped_lines)
                    
                    # Calculate space needed for text
                    text_height = num_lines * 0.045  # Space per line
                    text_padding = 0.08  # Space above and below text
                    total_text_space = text_height + text_padding
                    
                    # Axes position leaves room for text below
                    axes_bottom = 0.08 + total_text_space
                    axes_height = 0.96 - axes_bottom
                    
                    fig_height = 10.5  # Taller for text
                else:
                    # No description - standard layout
                    axes_bottom = 0.12
                    axes_height = 0.83
                    fig_height = 8.0
                
                # Create figure
                fig = plt.figure(figsize=(8, fig_height), dpi=350)
                ax = fig.add_axes([0.12, axes_bottom, 0.83, axes_height])
                ax.set_facecolor(BG_COLOR)
                fig.patch.set_facecolor(BG_COLOR)
                
                # Execute code
                exec_scope = {
                    "plt": plt, "np": np, "matplotlib": matplotlib,
                    "fig": fig, "ax": ax, "__builtins__": __builtins__
                }
                exec(task['code'], exec_scope)
                
                # ADD TEXT (NO BOX) if description exists
                if has_description:
                    # Calculate text position (centered below graph)
                    text_y = 0.08 + (total_text_space / 2)
                    
                    # Add sharp centered text WITHOUT box
                    fig.text(
                        0.50,              # x: centered
                        text_y,            # y: below graph
                        wrapped_text,
                        ha='center',       # horizontally centered
                        va='center',       # vertically centered
                        fontsize=15,       # Large for 50% readability (15pt → 7.5pt)
                        weight='bold',     # Bold for sharpness
                        color=CAPTION_RED, # Crimson red
                        family='DejaVu Sans',
                        linespacing=1.5,
                        transform=fig.transFigure
                    )
                
                # Save with maximum quality
                plt.savefig(img_name, facecolor=BG_COLOR, bbox_inches='tight',
                           pad_inches=0.3, dpi=350, format='png')
                plt.close('all')

                # Insert into document
                if task['para_index'] < len(doc.paragraphs):
                    target_para = doc.paragraphs[task['para_index']]
                    new_para = doc.add_paragraph()
                    target_para._p.addnext(new_para._p)
                    new_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    new_para.add_run().add_picture(img_name, width=Inches(6.0))
                    successful += 1
                    
                    if has_description:
                        print(f"✓ Para {task['para_index']}: Image with {num_lines}-line text")
                    else:
                        print(f"✓ Para {task['para_index']}: Image without text")
                
                if os.path.exists(img_name):
                    os.remove(img_name)
                    
            except Exception as e:
                failed += 1
                print(f"✗ Failed para {task['para_index']}: {str(e)[:120]}")
                plt.close('all')
                if os.path.exists(img_name):
                    os.remove(img_name)

        save_path = os.path.join(output_dir, 
                                 os.path.basename(doc_path).replace(".docx", "_Illustrated.docx"))
        doc.save(save_path)
        
        summary = f"""Successfully processed.
Saved: {save_path}
✓ Success: {successful} | ⚠ No description: {no_description} | ✗ Failed: {failed}"""
        print(summary)
        return summary

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AutoImageBuilderandPlacer - Sharp Text (No Box)")
        self.geometry("700x600")
        self.orchestrator = STEMOrchestrator()
        
        # Header
        header = ctk.CTkFrame(self, fg_color="#1e3a8a", corner_radius=10)
        header.pack(pady=20, padx=20, fill="x")
        
        ctk.CTkLabel(header, text="STEM Image Builder & Placer",
                    font=("Arial", 28, "bold"), text_color="white").pack(pady=15)
        
        ctk.CTkLabel(header, text="Sharp Text Below Graph - No Box",
                    font=("Arial", 11), text_color="#93c5fd").pack(pady=(0,15))
        
        # Features
        features = ctk.CTkFrame(self, fg_color="#f0f9ff", corner_radius=10)
        features.pack(pady=15, padx=20, fill="x")
        
        ctk.CTkLabel(features, text="✨ Features ✨",
                    font=("Arial", 13, "bold"), text_color="#1e3a8a").pack(pady=(10,5))
        
        for txt in [
            "📝 3-5 Line Descriptions Below Graph",
            "🎯 NO Box - Just Sharp Red Text",
            "🔍 15pt Bold - Readable at 50% (7.5pt)",
            "📐 Text Well Below Axes - Never Overlaps",
            "✅ DPI 350 - Ultra Sharp Quality"
        ]:
            ctk.CTkLabel(features, text=txt, font=("Arial", 10),
                        text_color="#374151").pack(pady=1)
        
        ctk.CTkLabel(self, text="DPI: 350 | Font: 15pt Bold DejaVu Sans | Model: gemini-3-flash-preview",
                    font=("Courier", 9), text_color="gray").pack(pady=8)
        
        # Buttons
        ctk.CTkButton(self, text="📄 Select Single File", command=self.run_file,
                     height=60, width=340, font=("Arial", 16, "bold"),
                     fg_color="#2563eb", hover_color="#1d4ed8").pack(pady=8)
        
        ctk.CTkButton(self, text="📁 Process Entire Folder", command=self.run_folder,
                     height=60, width=340, font=("Arial", 16, "bold"),
                     fg_color="#059669", hover_color="#047857").pack(pady=8)
        
        # Status
        status_frame = ctk.CTkFrame(self, fg_color="#ecfdf5", corner_radius=8)
        status_frame.pack(pady=10, padx=20, fill="x")
        
        self.status = ctk.CTkLabel(status_frame, text="● Ready - Text Without Box",
                                   text_color="#10b981", font=("Arial", 14, "bold"))
        self.status.pack(pady=10)
        
        # Console
        ctk.CTkLabel(self, text="📊 Console Output:",
                    font=("Arial", 12, "bold")).pack(pady=(10,0), padx=20, anchor="w")
        
        self.console = ctk.CTkTextbox(self, height=150, font=("Courier", 10),
                                      fg_color="#0f172a", text_color="#e2e8f0")
        self.console.pack(pady=5, padx=20, fill="both", expand=True)

    def log(self, msg):
        t = datetime.now().strftime("%H:%M:%S")
        self.console.insert("end", f"[{t}] {msg}\n")
        self.console.see("end")
        self.update()

    def run_file(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            self.status.configure(text="● AI Writing Descriptions...", text_color="#f59e0b")
            self.console.delete("1.0", "end")
            self.log("═" * 60)
            self.log(f"📄 FILE: {os.path.basename(path)}")
            self.log("═" * 60)
            self.update()
            
            result = self.orchestrator.process_file(path)
            self.log(f"\n{result}")
            self.log("═" * 60)
            messagebox.showinfo("✓ Complete", result)
            self.status.configure(text="● Ready - Text Without Box", text_color="#10b981")

    def run_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.status.configure(text="● Processing Folder...", text_color="#f59e0b")
            self.console.delete("1.0", "end")
            files = [f for f in os.listdir(folder) if f.endswith(".docx") and not f.startswith("~")]
            self.log("═" * 60)
            self.log(f"📁 FOLDER: {len(files)} documents")
            self.log("═" * 60 + "\n")
            
            for i, f in enumerate(files, 1):
                self.log(f"[{i}/{len(files)}] {f}")
                self.orchestrator.process_file(os.path.join(folder, f))
                self.log("")
            
            self.log("═" * 60)
            self.log(f"🎉 COMPLETE!")
            self.log("═" * 60)
            messagebox.showinfo("✓ Done", f"Processed {len(files)} files!")
            self.status.configure(text="● Ready - Text Without Box", text_color="#10b981")

if __name__ == "__main__":
    App().mainloop()