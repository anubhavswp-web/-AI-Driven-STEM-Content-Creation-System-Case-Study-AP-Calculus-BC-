# 🚀 Quick Start Guide

## Step 1: Install Python
- Download Python 3.10+ from [python.org](https://python.org)
- Make sure to check "Add Python to PATH" during installation

## Step 2: Install Required Libraries

Open Command Prompt (Windows) or Terminal (Mac/Linux) and run:

```bash
pip install google-genai python-docx matplotlib numpy customtkinter lxml
```

**Or use the requirements.txt file:**
```bash
pip install -r requirements.txt
```

## Step 3: Get Google Gemini API Key

1. Go to [Google AI Studio](https://aistudio.google.com/app/apikey)
2. Sign in with your Google account
3. Click "Create API Key"
4. Copy your API key

## Step 4: Create api_key.txt

In the same folder as the script, create a file named `api_key.txt` and paste your API key:

```
YOUR_API_KEY_HERE
```

## Step 5: Run the Script

```bash
python AutoImageBuilder_v12_ULTRA_HD.py
```

## Step 6: Use the GUI

1. Click "📄 Select Single File"
2. Choose your Word document
3. Wait for processing (images will be generated)
4. Find your output in the `Completed_Units/` folder

---

## ✅ Verification

To check if all libraries are installed:

```bash
python -c "import google.genai, docx, matplotlib, numpy, customtkinter, lxml; print('✓ All libraries installed!')"
```

---

## 📁 Folder Structure

```
your_project/
├── AutoImageBuilder_v12_ULTRA_HD.py    ← The script
├── api_key.txt                         ← Your API key
├── requirements.txt                    ← Library list (optional)
├── your_document.docx                  ← Input file
└── Completed_Units/                    ← Output folder (auto-created)
    └── your_document_Professional.docx ← Result!
```

---

## 🎯 Features You'll Get

- ✅ Ultra-HD 600 DPI images
- ✅ Pure black equations (#000000)
- ✅ 4-6 line detailed descriptions
- ✅ 15-20 professional visualizations per document
- ✅ Images inserted with descriptions

---

## ⚠️ Common Issues

### "pip is not recognized"
**Solution:** 
```bash
python -m pip install google-genai python-docx matplotlib numpy customtkinter lxml
```

### "No module named 'tkinter'"
**Solution (Linux):**
```bash
sudo apt-get install python3-tk
```

### Script won't run
**Solution:** Make sure you have Python 3.8 or newer:
```bash
python --version
```

---

## 💡 Tips

- Processing takes 3-5 seconds per image (15 images = ~1 minute)
- Larger documents take longer but produce better results
- Output files will be 15-30MB (high quality!)
- Keep your API key private (don't share `api_key.txt`)

---

**That's it! You're ready to create professional e-books with ultra-HD visualizations! 🎉**
