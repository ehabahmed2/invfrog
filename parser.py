"""
PDF Invoice Parser Module
Extracts text from PDFs and handles filename generation logic.
"""

import re
import os
import shutil
import pdfplumber
from datetime import datetime
from typing import Optional, Dict, Any, Tuple
from dataclasses import dataclass
from enum import Enum

class Status(Enum):
    OK = "OK"
    PARTIAL = "PARTIAL"
    SKIPPED = "SKIPPED"

class NamingScheme(Enum):
    INVOICE_NUMBER = "invoice_number"
    VENDOR_NAME = "vendor_name"
    ORIGINAL_FILENAME = "original_filename"

@dataclass
class ParseResult:
    filename: str
    status: Status
    reason: str
    data: Dict[str, Any]
    proposed_filename: str = ""
    target_path: str = ""

# Configuration
MAX_FILENAME_LENGTH = 200
ILLEGAL_CHARS = r'[\\/:*?"<>|]'

# --- Core Extraction Logic ---

def extract_text_from_pdf(filepath: str) -> Tuple[Optional[str], Optional[str]]:
    """Extracts text from first page. Returns (text, error_message)."""
    try:
        with pdfplumber.open(filepath) as pdf:
            if not pdf.pages:
                return None, "Empty PDF"
            text = pdf.pages[0].extract_text()
            if not text or len(text.strip()) < 20:
                return None, "Scanned PDF - unsupported"
            return text, None
    except Exception as e:
        err = str(e).lower()
        if 'password' in err: return None, "Encrypted/Password Protected"
        return None, f"Read Error: {str(e)[:50]}"

def parse_invoice_data(text: str) -> Dict[str, Any]:
    """Heuristic extraction of fields."""
    # Normalize
    text = re.sub(r'\r', '\n', text)
    
    data = {'invoice_number': None, 'date': None, 'total_amount': None, 'vendor': None}
    
    # 1. Invoice Number
    inv_patterns = [
        r'(?:Invoice\s*(?:No\.?|Number|#)\s*[:\s]*)\s*([A-Za-z0-9\-_/]{2,})',
        r'#\s*(\d{4,})'
    ]
    for p in inv_patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            if val.lower() not in ['date', 'no', 'number']:
                data['invoice_number'] = val
                break
    
    # 2. Date
    date_patterns = [
        r'(?:Date)[\s:]*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})',
        r'(\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2})',
        r'([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})'
    ]
    for p in date_patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            try:
                # Basic normalization attempt
                dstr = m.group(1).replace('/', '-').replace('.', '-')
                data['date'] = dstr 
                break
            except: pass

    # 3. Total Amount (Simplified Heuristic)
    # Look for "Total" followed by number, excluding "Subtotal"
    lines = text.split('\n')
    for line in lines:
        if 'total' in line.lower() and 'sub' not in line.lower():
            # Find last number in line usually
            amounts = re.findall(r'[\d,]+\.\d{2}', line)
            if amounts:
                data['total_amount'] = amounts[-1] # Take the last one (often the total)
                break
    
    # 4. Vendor (Very basic check for demonstration)
    vendor_keywords = ['inc', 'llc', 'ltd', 'gmbh', 'corp']
    for i, line in enumerate(lines[:5]): # Check header
        if any(k in line.lower() for k in vendor_keywords):
            data['vendor'] = line.strip()
            break

    return data

def process_single_pdf(filepath: str, filename: str) -> ParseResult:
    """Orchestrates extraction and returns result object."""
    text, error = extract_text_from_pdf(filepath)
    if error:
        return ParseResult(filename, Status.SKIPPED, error, {})
    
    data = parse_invoice_data(text)
    
    # Determine Status
    missing = [k for k, v in data.items() if not v and k != 'vendor']
    if not missing:
        status = Status.OK
        reason = "All fields found"
    elif len(missing) < 3:
        status = Status.PARTIAL
        reason = f"Missing: {', '.join(missing)}"
    else:
        status = Status.PARTIAL
        reason = "Low confidence extraction"
        
    return ParseResult(filename, status, reason, data)

# --- Organizer & Renaming Logic ---

def sanitize_filename(name: str) -> str:
    """Removes illegal characters."""
    if not name: return "unknown"
    name = re.sub(ILLEGAL_CHARS, '_', name)
    return name.strip()[:MAX_FILENAME_LENGTH]

def generate_proposed_filename(
    original_filename: str, 
    data: Dict[str, Any], 
    scheme: NamingScheme
) -> str:
    """Generates the new filename string (without path)."""
    
    # Helpers
    inv = sanitize_filename(data.get('invoice_number'))
    vendor = sanitize_filename(data.get('vendor') or "Unknown_Vendor")
    orig_base = os.path.splitext(original_filename)[0]
    
     # --- DATE FIX START ---
    # We want YYYYMMDD (e.g. 20120301) not 132012
    raw_date = data.get('date', '')
    date_formatted = ""
    
    if raw_date:
        try:
            # Try to parse the date string into a real date object
            # We try standard formats found in the Excel/CSV output
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
                try:
                    dt = datetime.strptime(raw_date, fmt)
                    date_formatted = dt.strftime("%Y%m%d") # Result: 20120301
                    break
                except ValueError:
                    continue
            
            # If standard parsing failed, just use the raw numbers but limit length
            if not date_formatted:
                nums = re.sub(r'[^0-9]', '', raw_date)
                if len(nums) >= 6: date_formatted = nums[:8]
                
        except:
            pass
    ext = ".pdf"
    
    # Helper to build name parts
    def build_name(*parts):
        # Filter out empty parts so we don't get double underscores
        return "_".join(filter(None, parts)) + ext  

    if scheme == NamingScheme.INVOICE_NUMBER:
        if inv and inv != "unknown":
            # Result: INV_47905_20120301.pdf
            return build_name("INV", inv, date_formatted)
        else:
            return build_name("INV_unknown", sanitize_filename(orig_base))
            
    elif scheme == NamingScheme.VENDOR_NAME:
        return build_name(vendor, "INV", inv, date_formatted)
        
    elif scheme == NamingScheme.ORIGINAL_FILENAME:
        if date_formatted:
            return f"{date_formatted}_{sanitize_filename(orig_base)}{ext}"
        return f"{sanitize_filename(orig_base)}{ext}"
    
    return original_filename

def calculate_target_path(
    base_output_folder: str,
    filename: str,
    data: Dict[str, Any],
    organize_by_month: bool
) -> str:
    """Calculates the full destination path."""
    
    subfolders = []
    
    # 1. Vendor Folder
    vendor = sanitize_filename(data.get('vendor') or "Unknown_Vendor")
    
    # 2. Date Folder (Optional)
    if organize_by_month:
        date_str = data.get('date', '')
        # Try to extract YYYY-MM. Crude heuristic for demo:
        # If date is 2023-01-01 or 01/01/2023
        year_month = "Unknown_Date"
        try:
            # Look for 4 digits
            m = re.search(r'20\d{2}', date_str)
            if m: year_month = m.group(0) # Just year for safety if parsing fails
        except: pass
        subfolders.append(year_month)
    
    subfolders.append(vendor)
    
    return os.path.join(base_output_folder, *subfolders, filename)

def get_safe_unique_path(target_path: str) -> str:
    """If file exists, appends _1, _2 etc."""
    if not os.path.exists(target_path):
        return target_path
        
    base, ext = os.path.splitext(target_path)
    counter = 1
    while True:
        new_path = f"{base}_{counter}{ext}"
        if not os.path.exists(new_path):
            return new_path
        counter += 1