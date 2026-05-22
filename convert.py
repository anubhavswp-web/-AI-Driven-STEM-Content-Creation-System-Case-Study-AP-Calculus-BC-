import subprocess
import os
import platform
import sys

def open_file(path):
    """Opens the Word doc immediately after creation."""
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", path])
        else:  # Linux
            subprocess.run(["xdg-open", path])
    except Exception as e:
        print(f"Could not open file automatically: {e}")

def run_conversion(text_content, output_name):
    """The engine that converts AI LaTeX text into a Word Doc."""
    temp_md = "temp_conversion_file.md"
    if not output_name.endswith(".docx"):
        output_name += ".docx"

    try:
        # Save the AI-generated text into a temporary Markdown file
        with open(temp_md, "w", encoding="utf-8") as f:
            f.write(text_content)
        
        # Pandoc command: 
        # Convert from Markdown to DOCX
        # Word natively handles the math conversion via Pandoc's internal filters
        subprocess.run(["pandoc", temp_md, "-o", output_name], check=True)
        
        print(f"\n✨ Success! Generated: {output_name}")
        
        # Cleanup and Auto-Open
        os.remove(temp_md)
        open_file(output_name)

    except FileNotFoundError:
        print("\n❌ Error: Pandoc is not installed on this system.")
        print("Download it here: https://pandoc.org/installing.html")
    except Exception as e:
        print(f"\n❌ Error during conversion: {e}")

def main():
    print("==========================================")
    print("   AI-to-Word Math Converter")
    print("==========================================")
    
    print("\nHow would you like to provide the AI output?")
    print("1) Paste the text directly")
    print("2) Load from a saved .txt or .md file")
    
    choice = input("\nSelect 1 or 2: ").strip()

    if choice == "1":
        print("\n--- Paste your AI content below ---")
        print("(When finished: Press Ctrl+Z then Enter on Windows, or Ctrl+D on Mac/Linux)")
        text_input = sys.stdin.read()
    elif choice == "2":
        file_path = input("Enter the filename (e.g., ai_output.txt): ").strip()
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as f:
                text_input = f.read()
        else:
            print("File not found!")
            return
    else:
        print("Invalid selection.")
        return

    out_file = input("\nWhat should we name the Word file? (e.g., Calculus_Homework): ").strip()
    if not out_file:
        out_file = "Converted_Math"

    run_conversion(text_input, out_file)

if __name__ == "__main__":
    main()