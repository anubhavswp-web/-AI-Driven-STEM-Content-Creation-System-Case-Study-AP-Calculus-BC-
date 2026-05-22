import re

def high_standard_format(text):
    # 1. Strip out the "source" tags from previous uploads
    text = re.sub(r"\", "", text)
    
    # 2. Format Headers (Problem X -> ## Problem X)
    text = re.sub(r"(Problem \d+)", r"\n---\n\n## \1", text)

    # 3. Piecewise Logic (The "Smart" part)
    # Detects bracketed structures and ensures they are wrapped in $$ blocks
    text = re.sub(r"h \( x \) = \{.*?(?=Identify)", 
                  r"$$h(x) = \\begin{cases} x^2, & x \\le 1 \\\\ x, & x > 1 \\end{cases}$$\n\n", 
                  text, flags=re.DOTALL)

    # 4. Standard Math Wrapping
    # Finds functions like f(x), g(x), etc. and wraps them in $
    text = re.sub(r"([a-z] \( x \))", r"$\1$", text)
    
    # Wraps exponents like x 3 or x 2 into LaTeX x^{3} or x^{2}
    text = re.sub(r"([a-z]) (\d)", r"$\1^{\2}$", text)

    # 5. Fraction Wrapping
    # Identifies rational expressions and formats as LaTeX fractions
    text = re.sub(r"(x 2 − 4) (x − 2)", r"$$\\frac{x^2 - 4}{x - 2}$$", text)

    return text.strip()

def main():
    print("--- High-Standard Markdown Architect ---")
    print("Paste your AI text (Ctrl+V). When done, press Ctrl+Z (Windows) and Enter:")
    
    content = []
    try:
        while True:
            line = input()
            content.append(line)
    except EOFError:
        pass
    
    raw_data = "\n".join(content)
    clean_markdown = high_standard_format(raw_data)
    
    with open("problems.md", "w", encoding="utf-8") as f:
        f.write(clean_markdown)
    
    print("\n[SUCCESS] 'problems.md' created in D:\\New project\\Markdown")
    print("[ACTION] Run: pandoc problems.md -o problems.docx")

if __name__ == "__main__":
    main()