# PDF Data Extraction Tool - User Guide FIRST DRAFT

A Python tool that automatically reads PDF utility bills and extracts important information like account numbers, dates, and charges into an organized Excel spreadsheet.

---

## What This Tool Does

This tool:
- Reads PDF files (like utility bills)
- Finds specific keywords (like "Amount Due", "Meter Number")
- Extracts the values associated with those keywords
- Saves everything to an Excel file with one sheet per PDF (Will adjust later) 
- Can run automatically every 24 hours to process new files

---

## Before You Start

### Install Required Software

#### 1. Python Packages
Open Terminal and run:
```bash
pip install pdfminer.six openpyxl pandas
```

- `pdfminer.six` - Reads text from PDFs
- `openpyxl` - Creates Excel files
- `pandas` - Organizes data into tables

#### 2. System Requirements
This tool works on Mac, Windows, and Linux.

---

## Setup

### 1. Create Your Folders

Create two folders in the same location as your Python script:

```
Your Project Folder/
├── testing_pdfminer.py          ← Your Python script
├── incoming_pdfs/                ← Put PDF files here
└── extracted_data/               ← Results will be saved here
```

To create folders:
```bash
mkdir incoming_pdfs
mkdir extracted_data
```

### 2. Configuring Keywords

The Python script has the `KEYWORD_OFFSETS` section at the bottom. This tells the tool what to look for:

```python
KEYWORD_OFFSETS = {
    'Amount Due': 2,           # Pulled value is 2 lines after "Amount Due"
    'Statement Date': 2,       # Value is 2 lines after "Statement Date"
    'Meter Number': 21,        # Value is 21 lines after "Meter Number"
    'kwh': 1                   # Extract only the number parts of the from "kwh" line
}
```

**Example:** If `'Amount Due': 2` means:
```
Line 50: Amount Due
Line 51: (blank line)
Line 52: $185.25  ← The tool will extract this (2 lines after)
```

---

## How to Use

### Option 1: Process PDFs Once (Testing)

1. Put your PDF files in the `incoming_pdfs/` folder
2. Run the script:
   ```bash
   python testing_pdfminer.py
   ```
3. Check the `extracted_data/` folder for results

### Option 2: Run Automatically Every 24 Hours

1. In your Python script, comment out Option 1 and uncomment Option 2:
   ```python
   # Option 1: Run once (for testing)
   # run_once(WATCH_FOLDER, OUTPUT_FOLDER, KEYWORD_OFFSETS)
   
   # Option 2: Run every 24 hours (continuous)
   run_continuously(WATCH_FOLDER, OUTPUT_FOLDER, KEYWORD_OFFSETS, interval_hours=24)
   ```

2. Run the script:
   ```bash
   python testing_pdfminer.py
   ```

3. Leave it running - it will automatically process new PDFs every 24 hours

4. To stop: Press `Ctrl + C`

---

## What You'll Get

### 1. Raw Text Files - I want to delete these later
For each PDF processed, you get a `.txt` file with all the extracted text:
```
extracted_data/
├── RGE_32_raw.txt          ← All text from the PDF
```

**Use these to:**
- Check if text was extracted correctly
- Find the right line offsets for your keywords
- Debug when values aren't being extracted
- Can be removed once code is fully implemented

### 2. Excel File with Results
One Excel file with all your data:
```
extracted_data/
└── extracted_results_20260320_143000.xlsx
    ├── Sheet: RGE_32       ← Data from first PDF
    ├── Sheet: RGE_33       ← Data from second PDF
    └── Sheet: RGE_34       ← Data from third PDF
```

**Each sheet contains:**
| keyword | found_at_line | line_offset | target_line | extracted_value | source_file |
|---------|---------------|-------------|-------------|-----------------|-------------|
| Amount Due | 50 | 2 | 52 | $185.25 | RGE_32.pdf |
| Meter Number | 25 | 21 | 46 | ABC123456 | RGE_32.pdf |

---

### Special Extraction Rules

The tool has built-in special handling for certain keywords:

#### Extract Numbers Only (for 'kwh' and 'ccf'):
```python
'kwh': 1    # From "Usage: 523 kwh" extracts → "523"
'ccf': 0    # From "45 ccf used" extracts → "45"
```

#### Extract Full Line (everything else):
```python
'Amount Due': 2           
'Statement Date': 2       
'Meter Number': 21        
```

---

## Understanding the Output

### Excel Columns Explained

- **keyword**: What word you searched for (e.g., "Amount Due")
- **found_at_line**: Which line number the keyword was found on
- **line_offset**: How many lines you told it to skip (from your settings)
- **target_line**: The actual line number where the value was extracted from
- **extracted_value**: The value that was extracted
- **match_type**: How the keyword was found ("single_line" or "split_lines")
- **source_file**: Which PDF file this came from
- **processed_date**: When the extraction happened

---

## Tips for Best Results

1. **Test with one PDF first** before processing many files
2. **Use the raw text files** to verify extraction and find correct offsets
3. **Keep your keywords specific** - "Total Amount Due" is better than just "Total"
4. **Process similar PDFs together** - PDFs from the same company usually have the same format
5. **Back up your original PDFs** before processing

---

## File Management

### Processed Files Log
The tool creates a `processed_files.json` file that tracks which PDFs have been processed. 

- **Don't delete this file** - it prevents processing the same PDF twice
- **Delete it to reprocess all PDFs** - if you want to start fresh

### Cleaning Up
To start fresh:
1. Delete `processed_files.json`
2. Empty the `extracted_data/` folder
3. Run the script again

---

## Version Information

- **Current Version:** 1.0
- **Last Updated:** March 2026
- **Compatible with:** Python 3.7+
