# K138 Form Filling System - User Guide

## Folder structure (Bhavesh-style)

- **templates** – Template PDFs and notice `.txt` files. Officers do not work in this folder.
- **working_case_example** – Put completed SAISIE PDFs here. The app creates the filled **K138 PDF in this folder**. No interim CSV files appear here.
- **Project root** – Scripts and (optionally) the executable. Interim CSVs are written only to a temp directory (e.g. `%LOCALAPPDATA%\RadianceCopilot\tmp`), not in templates or working case folders.

See **RUN.md** for run instructions and a quick check.

## What This System Does

This system automatically fills out K138 seizure notice forms by:
1. Reading data from completed SAISIE forms
2. Finding the right K138 template based on what you select
3. Automatically detecting where to put the text in the form
4. Loading legal notices from text files
5. Creating filled PDFs ready to use

---

## How to Use It

### Step 1: Set Up Your Templates Folder

Create a folder that contains:
- **SAISIE template PDF** - filename must have "saisie" and "template"
- **K138 template PDFs** - filenames must have "k138" and the form type (cannabis, knives, etc.)
- **Notice text files** (optional):
  - `k138_note_cannabis.txt` - Cannabis notice text
  - `k138_note_arms.txt` - Knives/Arms notice text
  - `k138_note_other.txt` - Other forms notice text (can be empty)

**Example templates folder contents:**
```
📁 Templates Folder
  ├── SAISIE À FAIRE_francompact 2025 - TEMPLATE.pdf
  ├── K138 Seizure Cannabis-Stupefiant - TEMPLATE-dummy.pdf
  ├── K138 Seizure Knives-Arms - TEMPLATE-dummy.pdf
  ├── K138 Stupefiant-Others - TEMPLATE-dummy.pdf
  ├── k138_note_cannabis.txt
  ├── k138_note_arms.txt
  └── k138_note_other.txt
```

**Note:** This folder is saved in `radiance_copilot.cfg` so you only need to select it once!

### Step 2: Run the Program

1. Open `saisie_a_faire_extractor.py`
2. **Select Templates Folder** - Browse to your templates folder (saved automatically)
3. **Select SAISIE PDF File** - Browse or drag-and-drop a filled SAISIE PDF file
4. **Choose Form Type** - Select Cannabis, Knives, or Others using radio buttons
5. **Click "Process PDF"** - The system processes the file and creates the filled K138 PDF

That's it! The system does everything else automatically.

---

## Features

### Persistent Settings
- **Templates folder** is saved in `radiance_copilot.cfg` - you only select it once
- **SAISIE PDF folder** is remembered - next time you browse, it opens to the last folder you used

### Drag-and-Drop Support
- **Drag PDF files** directly onto the "Select SAISIE PDF File" entry field
- Works with Windows Explorer and other file managers
- Automatically validates that it's a PDF file

### Notice Text Files
- Notice text is loaded from `.txt` files in the templates folder
- You can edit these files anytime without changing code
- If a notice file is missing, the form will be filled without that notice

---

## How It Finds Your PDFs

The system looks for keywords in the PDF filenames (it's not case-sensitive):

### Finding the SAISIE Template
- **Looks for**: PDFs in templates folder that have BOTH "saisie" AND "template" in the name
- **Example**: `SAISIE À FAIRE_francompact 2025 - TEMPLATE.pdf` ✓

### Finding the K138 Template
- **Looks for**: PDFs in templates folder that have "k138" AND a keyword matching your form type:
  - **Cannabis**: looks for "cannabis"
  - **Knives**: looks for "knife", "knives", or "arms"
  - **Others**: looks for "stupefiant" (but not cannabis or knives)
- **Examples**:
  - Cannabis form → `K138 Seizure Cannabis-Stupefiant - TEMPLATE-dummy.pdf` ✓
  - Knives form → `K138 Seizure Knives-Arms - TEMPLATE-dummy.pdf` ✓
  - Others form → `K138 Stupefiant-Others - TEMPLATE-dummy.pdf` ✓

### Finding Notice Text Files
- **Looks for**: Text files in templates folder:
  - `k138_note_cannabis.txt` for Cannabis forms
  - `k138_note_arms.txt` for Knives/Arms forms
  - `k138_note_other.txt` for Other forms
- If a file is missing, no notice is added (form still processes successfully)

---

## How the Box Gets Filled Out

### The Problem We Solved

Before: You had to manually figure out where the text box was in the PDF and type in exact coordinates. If the template changed, you had to do it all over again.

Now: The system automatically finds the text box for you!

### How It Works

1. **Opens the K138 PDF** and looks inside it
2. **Finds the description box** - It's looking for a form field (usually named "C")
3. **Reads the box location** - Gets the exact position:
   - How far from the left edge
   - How far from the right edge
   - How tall the box is
   - Where it starts and ends
4. **Uses those measurements** to make sure all text fits perfectly inside

### Why This Is Important

- ✅ Text always fits in the box (no overflow)
- ✅ Works with any K138 template automatically
- ✅ No manual adjustments needed
- ✅ If you get a new template, it just works

### How Text Gets Wrapped

The system measures each line of text to make sure it fits:

1. **Takes your text** (including notice text from .txt files)
2. **Measures how wide it is** using the actual font size
3. **Breaks it into lines** that fit within the box width
4. **Wraps at word boundaries** (doesn't cut words in half)
5. **Handles long URLs** by splitting them if needed

**Example:**
```
Long text: "Please note: Under the Cannabis Act, the above item is illegal..."

Gets wrapped to:
Line 1: "Please note: Under the Cannabis Act, the above item is"
Line 2: "illegal to transport, import or export across Canada's"
Line 3: "borders, including by mail or courier..."
```

---

## What Happens When You Click "Process PDF"

### Step 1: Finding Files
- ✅ Finds the SAISIE template in templates folder
- ✅ Finds the matching K138 template (based on your form type selection)
- ✅ Loads notice text from the appropriate .txt file (if it exists)

### Step 2: Processing the PDF

1. **Extracts the data**:
   - Compares the completed SAISIE PDF with the template
   - Finds what was filled in (address, date, item description, etc.)
   - Saves it to a CSV file (same folder as the SAISIE PDF)

2. **Prepares K138 data**:
   - Formats the extracted data for the K138 form
   - Adds the form type you selected
   - Adds notice text from the .txt file (if available)
   - Creates a `k138_values.csv` file (same folder as the SAISIE PDF)

3. **Fills the K138 form**:
   - Opens the K138 template
   - Detects the box dimensions automatically
   - Fills in all the data
   - Includes the notice text from the .txt file (if loaded)
   - Saves the filled PDF as `K138_filled_from_csv.pdf` (same folder as the SAISIE PDF)

### Step 3: Summary
Shows you a success message and logs all actions in the log area.

---

## Notice Text Files

### File Format
- **Plain text files** (.txt) in the templates folder
- **Can contain** any text you want to appear at the bottom of the description box
- **Multi-line text** is supported (use line breaks)

### Example: k138_note_cannabis.txt
```
(Please note: Under the Cannabis Act, the above item is illegal to transport, import or export across Canada's borders, including by mail or courier. For more information please visit this website: https://www.canada.ca/en/services/health/campaigns/cannabis/border.html)
-------------------------------------------------------------------------------------------------------
(Attention : La Loi sur le cannabis interdit d'importer l'article ci-dessus au Canada ou de l'en exporter, y compris par la poste ou par messagerie. Pour en savoir plus : https://www.canada.ca/fr/services/sante/campagnes/cannabis/frontiere.html)
```

### Example: k138_note_arms.txt
```
CENTRIFUGAL KNIFE AS PER MEMORANDUM D19-13-2
COUTEAU CENTRIFUGE AU MÉMORANDUM D19-13-2
```

### Example: k138_note_other.txt
```
(Can be empty or contain any notice text for other form types)
```

**Important:** 
- If a notice file is missing, the form will still process successfully (just without that notice)
- You can edit these files anytime - no code changes needed
- The notice text is added to the description item field automatically

---

## Configuration File

The system saves your preferences in `radiance_copilot.cfg`:

```
[paths]
templates_folder = C:\Users\...\Templates Folder
saisie_folder = C:\Users\...\saisie aeads\folder123
```

- **templates_folder**: Your templates folder path (saved when you select it)
- **saisie_folder**: Last folder you selected a SAISIE PDF from (for convenience)

These paths are loaded automatically when you start the program.

---

## Troubleshooting

### "Permission denied" Error
**What it means**: A file is open in another program (like Excel or a PDF viewer)

**What to do**:
1. Close any CSV or PDF files that might be open
2. Wait a moment if OneDrive is syncing
3. Try again

### "Could not find PDF" Error
**What it means**: The system couldn't find a PDF with the right keywords

**What to check**:
1. Make sure your PDF filenames contain the required keywords
2. Make sure you selected the correct templates folder
3. Make sure the K138 template matches your selected form type

### "No notice text file found" Warning
**What it means**: The system couldn't find the .txt file for the notice

**What to do**:
1. Create the missing .txt file in your templates folder
2. Or ignore it - the form will process without the notice
3. The form will still work correctly

### Drag-and-Drop Not Working
**What it means**: The tkinterdnd2 library might not be installed

**What to do**:
1. Install it: `pip install tkinterdnd2`
2. Or just use the "Browse" button instead

---

## What Was Added to the System

### Before (Old System):
- ❌ Had to drag and drop each file individually
- ❌ Had to manually figure out box coordinates
- ❌ Had to process one PDF at a time
- ❌ Had to guess the form type
- ❌ Had to select folders every time
- ❌ Notice text was hardcoded in the program

### After (New System):
- ✅ **Templates folder** - Select once, saved automatically
- ✅ **Individual file processing** - Select or drag-and-drop one SAISIE PDF at a time
- ✅ **Automatic box detection** - System finds the box location automatically
- ✅ **Form type selection** - Choose with radio buttons
- ✅ **Smart text wrapping** - Text always fits perfectly
- ✅ **Notice text from files** - Edit .txt files without changing code
- ✅ **Config file persistence** - Remembers your folder selections
- ✅ **Drag-and-drop support** - Drop PDF files directly onto the entry field
- ✅ **Better error handling** - Continues even if one PDF has issues

---

## Technical Summary (For Developers)

### Config File Handling
- Uses `configparser` to read/write `radiance_copilot.cfg`
- Saves templates folder and last SAISIE folder path
- Loads paths on startup automatically

### Notice Text Loading
- Reads from `.txt` files in templates folder: `k138_note_cannabis.txt`, `k138_note_arms.txt`, `k138_note_other.txt`
- Adds notice text to `description_item` field when building K138 values
- Notice is included in `description_block` automatically via CSV reconstruction

### Box Detection
- Uses PyMuPDF to scan PDF form fields
- Finds field "C" (description box)
- Extracts coordinates and converts from PyMuPDF to ReportLab coordinate system
- Formula: `reportlab_y = page_height - pymupdf_y`

### Text Wrapping
- Uses ReportLab's `stringWidth()` to measure actual text width
- Wraps at word boundaries
- Handles long words/URLs by character splitting if needed

### PDF Search
- Case-insensitive keyword matching in filenames
- Form type keywords: cannabis, knives/arms, stupefiant
- Searches in templates folder (not main folder)

### File Processing
- Processes individual SAISIE PDF files (not batch)
- **Output**: K138 PDF saved to working case folder (same folder as input SAISIE PDF)
- **Interim files**: CSVs stored in AppData/temp, auto-deleted when done
- Output filename: `K138_{input_stem}.pdf` (e.g. `K138_20260129103247504.pdf`) — avoids overwriting when multiple seizures in same folder
- Templates folder is read-only; never receives generated outputs

---

## How to Test

1. Put a Saisie PDF in `working_case_example/` (or any working case folder).
2. Run `saisie_a_faire_extractor.py` (or R-Copilot.exe).
3. Select templates folder, select the Saisie PDF, choose form type, click Process PDF.
4. **Confirm**: `K138_{saisie_stem}.pdf` appears in `working_case_example/`.
5. **Confirm**: No CSV or temp files appear in `working_case_example/` or `templates/`.
6. **Sanity check**: If you select a Saisie PDF from inside the templates folder, the app should show an error ("Wrong File").

---

## Quick Start Checklist

- [ ] Create templates folder with template PDFs
- [ ] Create notice text files (k138_note_*.txt) in templates folder
- [ ] Run `saisie_a_faire_extractor.py`
- [ ] Select templates folder (saved automatically)
- [ ] Select or drag-and-drop a SAISIE PDF file
- [ ] Choose form type
- [ ] Click "Process PDF"
- [ ] Done! Check the output folder for `K138_{saisie_stem}.pdf`

---

## Dependencies

Install these Python packages:
```bash
pip install PyPDF2 reportlab PyMuPDF tkinterdnd2
```

**Note:** `tkinterdnd2` is optional but recommended for drag-and-drop support.

That's it! The system handles everything else automatically.
