# PDF Bill Scanner

An automated data extraction pipeline for energy utility bills. Scans a folder of PDF bills, extracts key fields, and appends results as rows to an existing Excel workbook.

---

## Requirements

- Python 3.x — download from [python.org](https://www.python.org/) if not already installed
- Microsoft Excel or compatible spreadsheet software
- The following files must all be in the same folder:
  - `main.py`
  - `runbutton.py`
  - Your target Excel workbook (e.g. `Staging Document.xlsx`)

---

## First-Time Setup

### 1. Open a terminal and navigate to the project folder

**Mac:**
```bash
cd /path/to/your/project/folder
```

**Windows:**
```bash
cd C:\path\to\your\project\folder
```

---

### 2. Create a virtual environment

```bash
python -m venv .venv
```

---

### 3. Activate the virtual environment

**Mac:**
```bash
source .venv/bin/activate
```

**Windows:**
```bash
.venv\Scripts\activate
```

You will know it worked when `(.venv)` appears at the start of your terminal line.

---

### 4. Install required packages

```bash
pip install pdfminer.six pandas openpyxl
```

---

### 5. Create the required folders

Create two folders inside the project folder if they do not already exist:

- `incoming_pdfs` — place PDF bills here before running a scan
- `extracted_data` — the program saves debug files here automatically

**Mac:**
```bash
mkdir incoming_pdfs extracted_data
```

**Windows:**
```bash
mkdir incoming_pdfs
mkdir extracted_data
```

---

### 6. Prepare the Excel workbook

The program appends to an **existing** Excel file — it will not create one from scratch. Before running for the first time:

- Open your Excel workbook
- Make sure the target sheet exists and its name **exactly matches** the sheet name defined in `main.py` (default: `"Input Data"`)
- Make sure at least one sheet in the workbook is **visible** (not hidden)
- Save and close the file before running the program

---

### 7. Set the Excel file path in `main.py`

Open `main.py` and update the `EXCEL_FILE` variable at the bottom to point to your workbook:

```python
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
WATCH_FOLDER  = os.path.join(BASE_DIR, 'incoming_pdfs')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'extracted_data')
EXCEL_FILE    = os.path.join(BASE_DIR, 'YourWorkbookName.xlsx')
```

> **Important:** Using `os.path.join(BASE_DIR, ...)` as shown above is preferable over a hardcoded absolute path. Defining a base directory allows for the main location to be the starting point, ensuring the program works correctly on both Mac and Windows without modification. Hardcoding the path works when system ownership and operation is controlled by one device only. 

---

## Running the Program

### Option A — Using the button UI (recommended)

With the virtual environment activated, run:

```bash
python run_button.py
```

A window will appear with a **RUN SCAN** button. Press it to trigger a scan. Output from the scan streams live into the log panel in the window. The button will turn green on success or red on error.

> The UI automatically detects and uses the `.venv` Python environment — you do not need to activate it manually each time as long as you launch via `run_button.py`.

### Option B — Running directly from the terminal

With the virtual environment activated, run:

```bash
python main.py
```

---

## How to Run a Scan

1. Place one or more PDF energy bills into the `incoming_pdfs` folder
2. Make sure the target Excel file is **closed** in Excel before scanning
3. Press **RUN SCAN** (or run `main.py` directly)
4. Results are appended as new rows to the target Excel sheet
5. Debug text files for each PDF are saved to `extracted_data`

---

## File Structure

```
project-folder/
│
├── main.py                  # Main extraction pipeline
├── run_button.py            # Button UI launcher
├── processed_files.json     # Auto-generated log of processed PDFs
│
├── incoming_pdfs/           # Place new PDF bills here
├── extracted_data/          # Debug output (_raw.txt files)
│
├── YourWorkbook.xlsx        # Target Excel file (you provide this)
└── .venv/                   # Virtual environment (created during setup)
```

---

## Common Errors and Fixes

| Error | Cause | Fix |
|-------|-------|-----|
| `ModuleNotFoundError: 'pdfminer' is not a package` | A file in the folder is named `pdfminer.py` | Rename the file to anything else |
| `ImportError` on `pdfminer`, `pandas`, or `openpyxl` | Virtual environment not active or packages not installed | Activate `.venv` and run `pip install pdfminer.six pandas openpyxl` |
| `IndexError: At least one sheet must be visible` | All sheets in the Excel file are hidden | Open the workbook, right-click a sheet tab, select Unhide |
| `Sheet 'X' not found` | Sheet name in code does not match the actual tab name | Open `main.py` and update the `sheet_name` variable to match exactly |
| `PermissionError` creating `extracted_data` | Script is running from a protected folder (e.g. OneDrive) | Ensure `WATCH_FOLDER` and `OUTPUT_FOLDER` use `os.path.join(BASE_DIR, ...)` |
| `No new files found` on every run | PDFs are already in the processed log | Delete `processed_files.json` to force a full re-scan |
| `No new files found` — files not detected | PDFs are in the wrong folder | Confirm PDFs are inside the `incoming_pdfs` folder |
| Blank or wrong values extracted | Keyword line offsets are misaligned for this bill format | Open the `_raw.txt` debug file and count lines manually to find the correct offset |
| Excel file not found | `EXCEL_FILE` path is wrong or file does not exist | Confirm the workbook exists and the path in `main.py` is correct |
| `processed_files.json` not recognizing files on a new machine | Log stores absolute paths from the old machine | Delete `processed_files.json` — it will rebuild on the next run |

---

## Moving to a New Device

1. Copy the entire project folder to the new device
2. Delete `processed_files.json` if present — paths from the old machine will not match
3. Follow the **First-Time Setup** steps above from step 1
4. Update `EXCEL_FILE` in `main.py` if the workbook location has changed

> Do **not** copy the `.venv` folder between machines — virtual environments are not portable. Always create a fresh one on each device using the setup steps above.

---

## Notes

- **Close the Excel workbook before running a scan** — openpyxl cannot write to a file that Excel has open
- The `extracted_data` folder contains a `_raw.txt` file for each processed PDF showing the full text as pdfminer sees it — useful for debugging incorrect extractions
- Keyword offsets in `KEYWORD_OFFSETS` were calibrated against RGE bill format 32 — if bills from a different utility are introduced, offsets may need adjustment
- The program will not re-process a PDF it has already seen unless `processed_files.json` is deleted or the PDF file is modified
