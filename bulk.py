import os
import time
import json
from datetime import datetime
from pathlib import Path
import pandas as pd
from pdfminer.high_level import extract_text
from pdfminer.layout import LAParams
import re
from openpyxl import load_workbook, Workbook

# PDF MINER FUNCTION - SWITCH TO OCR IF NECESSARY
def extract_pdf_text(pdf_path):
    """
    Extract text from PDF using pdfminer (NOT OCR YET).
    If we change this to Microsoft Azure OCR, replace this section (# PDF MINER FUNCTION to #END PDF MINER FUNCTION) with new Azure commands. 

    Args:
        pdf_path: Path to PDF file
    
    Returns:
        Extracted text as string
    """
    # Configure layout analysis parameters for better table extraction
    laparams = LAParams(
        line_margin=0.5,
        char_margin=2.0,
        word_margin=10.0,  # Large to capture right-aligned if possible
        boxes_flow=0.5,
        detect_vertical=False,
        all_texts=True
    )
    
    text = extract_text(pdf_path, laparams=laparams)
    return text

# END PDF MINER FUNCTION


# KEYWORD SEARCH

def extract_values_after_keywords(pdf_path, keyword_offset_dict):
    """
    Extract values X lines after keywords from the pdf mined text.
    Each keyword can have its own specific line offset to account for PDF irreguarity.
    Can handle keywords split across two lines with method 2.
    
    Args:
        pdf_path: Path to PDF file
        keyword_offset_dict: Dictionary mapping keywords to their line offsets
            Example: {'Date': 2, 'Amount Due': 2, 'Meter Number': 21}
    
    Returns:
        DataFrame with all extraction results - Subsequent edits for excel exporting soon.
    """
    text = extract_pdf_text(pdf_path)
    lines = text.split('\n')
    results = []
    integer_pattern = r'\b\d+\b'

    for keyword, lines_offset in keyword_offset_dict.items():
        # Remove spaces from keyword for matching
        keyword_no_spaces = keyword.replace(' ', '')
        pattern = re.compile(re.escape(keyword), re.IGNORECASE)
        
        # Method 1: Search for keyword on single line (normal case)
        for line_num, line in enumerate(lines):
            match = pattern.search(line)
            if match:
                target_line_num = line_num + lines_offset
                if target_line_num < len(lines):
                    target_line = lines[target_line_num].strip()
                    
                    # SPECIAL HANDLING for specific keywords
                    if keyword.lower() in ['ccf', 'kwh']:
                        # Extract integer only
                        int_match = re.search(integer_pattern, target_line)
                        extracted_value = int_match.group() if int_match else ""
                    else:
                        # For ALL other keywords: extract entire line
                        extracted_value = target_line
                else:
                    extracted_value = ""
                
                # Only add if extracted_value is not blank
                if extracted_value:
                    results.append({
                        'keyword': keyword,
                        'found_at_line': line_num,
                        'line_offset': lines_offset,
                        'target_line': target_line_num,
                        'extracted_value': extracted_value,
                        'match_type': 'single_line'
                    })
            
    # Method 2: Search for keyword split across two consecutive lines
        # This is the "meter number" method (the only way we can pull meter number) and is not functioning as well as it used to? investigate
    for line_num in range(len(lines) - 1):
        # Combine current line and next line (removing spaces/newlines)
        combined = (lines[line_num] + lines[line_num + 1]).replace(' ', '').replace('\n', '')
        
        # Check if keyword (without spaces) matches the combined lines
        if keyword_no_spaces.lower() in combined.lower():
            # The keyword ends on line_num + 1, so offset starts from there
            target_line_num = (line_num + 1) + lines_offset
            
            if target_line_num < len(lines):
                target_line = lines[target_line_num].strip()
                
                # SPECIAL HANDLING for specific keywords
                if keyword.lower() in ['ccf', 'kwh']:
                    # Extract integer only
                    int_match = re.search(integer_pattern, target_line)
                    extracted_value = int_match.group() if int_match else ""
                else:
                    # For ALL other keywords: extract entire line
                    extracted_value = target_line
            else:
                extracted_value = ""
            
            # Only add if extracted_value is not blank
            if extracted_value:
                results.append({
                    'keyword': keyword,
                    'found_at_line': f"{line_num}-{line_num+1} (split)",
                    'line_offset': lines_offset,
                    'target_line': target_line_num,
                    'extracted_value': extracted_value,
                    'match_type': 'split_lines'
                })
        
    df = pd.DataFrame(results)
    # Remove duplicate matches (prefer single line matches over split matches)
    if not df.empty:
        df = df.drop_duplicates(subset=['keyword', 'extracted_value'], keep='first')
    
    return df

# END KEYWORD SEARCH


# AUTOMATED PDF SCANNING & EXPORTING DATA + INTERMEDIATES FOR CHECKS
    # When implementing final, delete the _raw.txt files for cleanliness 

class PDFScanner:
    def __init__(self, watch_folder, output_folder, processed_log='processed_files.json'):
        """
        Scanning!
        
        Args:
            watch_folder: Folder to monitor for new PDFs
            output_folder: Folder to save extraction results
            processed_log: JSON file to track processed files
        """
        self.watch_folder = Path(watch_folder)
        self.output_folder = Path(output_folder)
        self.processed_log = processed_log
        
        # Create folders if they don't exist
        self.output_folder.mkdir(exist_ok=True)
        
        # Load or create processed files log
        self.processed_files = self.load_processed_log()
    
    def load_processed_log(self):
        # "Load the list of already processed files."
        if os.path.exists(self.processed_log):
            with open(self.processed_log, 'r') as f:
                return json.load(f)
        return {}
    
    def save_processed_log(self):
        # "Save the updated list of processed files."
        with open(self.processed_log, 'w') as f:
            json.dump(self.processed_files, f, indent=2)
    
    def get_new_files(self):
        # "Find PDF files that haven't been processed yet." 
        # Currently searching by name. Update to metadata (upload time, etc.) if necessary
        all_pdfs = list(self.watch_folder.glob('*.pdf'))
        new_files = []
        
        for pdf_file in all_pdfs:
            file_key = str(pdf_file)
            file_modified = os.path.getmtime(pdf_file)
            
            # Check if file is new or has been modified
            if file_key not in self.processed_files or \
               self.processed_files[file_key] != file_modified:
                new_files.append(pdf_file)
        
        return new_files
    
    def process_single_file(self, pdf_path, keyword_offset_dict):
        """
        Process a single PDF file and extract data.
        
        Args:
            pdf_path: Path to PDF file
            keyword_offset_dict: Dictionary mapping keywords to line offsets
        
        Returns:
            DataFrame with extraction results
        """
        print(f"\n{'='*60}")
        print(f"Processing: {pdf_path.name}")
        print(f"{'='*60}")
        
        try:
            # Extract full text using pdfminer
            full_text = extract_pdf_text(str(pdf_path))
            
            # Save the raw txt output to a text file - for debugging, delete in final implementation
            txt_output_file = self.output_folder / f"{pdf_path.stem}_raw.txt"
            with open(txt_output_file, 'w', encoding='utf-8') as f:
                f.write(full_text)
            print(f"✓ Saved raw text to {txt_output_file}")
            
            # Extract values using keyword-specific offsets
            df = extract_values_after_keywords(str(pdf_path), keyword_offset_dict)
            
            # Add metadata
            df['source_file'] = pdf_path.name
            df['processed_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            print(f"✓ Extracted {len(df)} values")
            return df
            
        except Exception as e:
            print(f"✗ Error processing {pdf_path.name}: {e}")
            return None
    
    def scan_and_process(self, keyword_offset_dict):
        """
        Scan for new files and process them.
        Saves results to Excel with each PDF on its own sheet.
        """
        print(f"\n{'#'*60}")
        print(f"Starting scan at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Watching folder: {self.watch_folder}")
        print(f"{'#'*60}")
        
        # Find new files
        new_files = self.get_new_files()
        
        if not new_files:
            print("✓ No new files found")
            return
        
        print(f"✓ Found {len(new_files)} new file(s) to process")
        
        # Create new Excel file path for debugging - Change this in future. All results to one running excel 
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_file = self.output_folder / f"extracted_results_{timestamp}.xlsx"
        
        # Flag to track first sheet
        first_sheet = True
        
        # Process each new file
        for pdf_file in new_files:
            df = self.process_single_file(pdf_file, keyword_offset_dict)
            
            if df is not None and not df.empty:
                sheet_name = pdf_file.stem[:31]
                
                # Write to Excel - Change this in future. All results to one running excel 
                if first_sheet:
                    with pd.ExcelWriter(excel_file, engine="openpyxl", mode="w") as writer:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                    first_sheet = False
                    print(f"✓ Created Excel file with sheet '{sheet_name}'")
                else:
                    with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"✓ Added sheet '{sheet_name}' to Excel file")
                
                # Mark file as processed
                self.processed_files[str(pdf_file)] = os.path.getmtime(pdf_file)
        
        # Update processed files log
        self.save_processed_log()
        print(f"\n✓ All results saved to {excel_file}")
        print(f"✓ Scan complete!")


def run_once(watch_folder, output_folder, keyword_offset_dict):
    """Run the scanner once."""
    scanner = PDFScanner(watch_folder, output_folder)
    scanner.scan_and_process(keyword_offset_dict)

def run_continuously(watch_folder, output_folder, keyword_offset_dict, interval_hours=24):
    """
    Run the scanner continuously at specified intervals.
    
    Args:
        watch_folder: Folder to monitor for PDFs
        output_folder: Folder to save results
        keyword_offset_dict: Dictionary mapping keywords to line offsets
        interval_hours: How often to scan (default: 24 hours)
    """
    scanner = PDFScanner(watch_folder, output_folder)
    
    print(f"\n{'#'*60}")
    print(f"AUTOMATED PDF SCANNER STARTED")
    print(f"Scanning every 24 hour(s)")
    print(f"Press Ctrl+C to stop")
    print(f"{'#'*60}\n")
    
    try:
        while True:
            scanner.scan_and_process(keyword_offset_dict)
            
            # Wait for next scan
            wait_seconds = interval_hours * 3600
            next_scan = datetime.now().timestamp() + wait_seconds
            next_scan_time = datetime.fromtimestamp(next_scan).strftime('%Y-%m-%d %H:%M:%S')
            
            print(f"\n{'='*60}")
            print(f"Sleeping until next scan at {next_scan_time}")
            print(f"{'='*60}\n")
            
            time.sleep(wait_seconds)
            
    except KeyboardInterrupt:
        print("\n\n✓ Scanner stopped by user")

# END CONT. SCANNING


# USAGE 
# KEYWORD DEFINITIONS AND RUN OPTIONS

if __name__ == "__main__":
    # *** NAME FOLDERS AS WE NEED ***
    WATCH_FOLDER = 'incoming_pdfs'  # <<< Folder to monitor for new PDFs
    OUTPUT_FOLDER = 'extracted_data'  # <<< Where to save results
    
    # HARDCODED KEYWORDS AND LINE OFFSETS 
    # Using RGE 32 as refrence
    KEYWORD_OFFSETS = {
    'Amount Due': 2,                    
    'Statement Date': 2,                
    'Meter Number': 22,                           
    'Total Natural Gas Cost': 15,       
    'Total Electricity Cost': 38,       
    'ccf': 0,                           # Special export: integer only
    'kwh': 0                            # Special export: integer only
}
    
    # RUN OPTIONS for testing vs actual implementation
    
    # Option 1: Run once (for testing)
    run_once(WATCH_FOLDER, OUTPUT_FOLDER, KEYWORD_OFFSETS)
    
    # Option 2: Run every 24 hours (continuous, final product)
    # run_continuously(WATCH_FOLDER, OUTPUT_FOLDER, KEYWORD_OFFSETS, interval_hours=24)
    
    # Option 3: Run every hour (for testing)
    # run_continuously(WATCH_FOLDER, OUTPUT_FOLDER, KEYWORD_OFFSETS, interval_hours=1)

# END :) 
