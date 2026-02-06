================================================================================
                              InvFrog
                    PDF Invoice Extractor & Organizer
                              User Guide
================================================================================

WHAT THIS TOOL DOES
-------------------
InvFrog extracts key data from PDF invoices and saves it to an Excel file.
It can also rename and organize your invoice files into structured folders.

Fields extracted:
  - Invoice Number
  - Invoice Date
  - Total Amount
  - Vendor/Company Name (when explicitly labeled)

New Organize & Rename Feature:
  - Renames PDFs based on extracted invoice data
  - Organizes files into folders by vendor and/or date
  - Copies files (never moves originals) for safety
  - Dry Run mode to preview changes before applying

IMPORTANT LIMITATIONS
---------------------
* Works ONLY with machine-generated PDFs (digital invoices)
* Does NOT work with scanned documents or images of paper invoices
* No OCR capability - text must be selectable in the PDF

================================================================================
                           HOW TO USE
================================================================================

BASIC WORKFLOW
--------------
1. Double-click the application to open it
2. Click "Browse..." to select input folder containing PDF invoices
3. (Optional) Select a different output folder
4. Configure settings:
   - Choose naming scheme
   - Enable/disable "Organize by Year-Month"
   - Enable "Dry Run" to preview first (RECOMMENDED)
5. Click "Start Extraction & Organize"
6. Review results in the table
7. If satisfied with preview, disable "Dry Run" and run again

DRY RUN MODE (Recommended First Step)
-------------------------------------
Dry Run lets you preview exactly what will happen WITHOUT copying any files.

When enabled:
  - Shows proposed new filenames in the table
  - Creates preview.csv with all planned changes
  - Does NOT copy or rename any files
  - Perfect for verifying naming scheme works for your invoices

After reviewing:
  - If changes look correct: Uncheck "Dry Run" and run again
  - If changes need adjustment: Modify settings and preview again

================================================================================
                         NAMING SCHEMES
================================================================================

1. Invoice Number (Recommended)
   Format: INV_{InvoiceNumber}_{YYYYMMDD}.pdf
   Example: INV_12345_20240115.pdf
   If no invoice number: INV_unknown_{original_filename}.pdf

2. Vendor Name
   Format: {Vendor}_INV_{InvoiceNumber}_{YYYYMMDD}.pdf
   Example: AcmeCorp_INV_12345_20240115.pdf
   If no vendor: Unknown_INV_{number}_{date}.pdf

3. Original Filename
   Format: {YYYYMMDD}_{original_filename}.pdf
   Example: 20240115_supplier_invoice.pdf
   Keeps original name but adds date prefix

================================================================================
                       FOLDER ORGANIZATION
================================================================================

When "Organize into folders by Year-Month" is CHECKED:
  output_folder/
    2024-01/
      Vendor_Name/
        INV_12345_20240115.pdf
    2024-02/
      Vendor_Name/
        INV_12346_20240210.pdf
    Unknown_Vendor/
      INV_unknown_miscfile.pdf

When UNCHECKED:
  output_folder/
    Vendor_Name/
      INV_12345_20240115.pdf
    Unknown_Vendor/
      INV_unknown_miscfile.pdf

================================================================================
                        OUTPUT FILES
================================================================================

After processing, you'll find these files in your output folder:

* Invoices_Extracted_{timestamp}.xlsx - All extracted invoice data
* preview.csv - (Dry Run only) Planned changes
* skipped_files.csv - List of PDFs that couldn't be processed
* errors.log - Technical details for troubleshooting

IMPORTANT: Files are COPIED, not moved. Your original files remain unchanged.

================================================================================
                       STATUS MEANINGS
================================================================================

* OK      - All main fields extracted successfully
* PARTIAL - Some fields missing but file was processed
* SKIPPED - File could not be processed (see reason)

COMMON SKIP REASONS
-------------------
* "Scanned / Image PDF" - PDF contains images, not text
* "Protected / Encrypted PDF" - PDF requires a password
* "Corrupted PDF file" - PDF file is damaged
* "Multiple totals detected" - Ambiguous amounts found
* "Permission denied" - Cannot read/write file

================================================================================
                    FIELD EXTRACTION RULES
================================================================================

Total Amount:
  - Looks for: "Invoice Total", "Total Due", "Balance Due", "Amount Due"
  - Excludes lines with: Wire, A/R, Reference, Net, Contract, Terms
  - If multiple different totals: marked as PARTIAL

Vendor/Company Name:
  - ONLY extracted if PDF contains explicit labels:
    * "Company:", "Vendor:", "From:", "Seller:", "Remit To:"
  - Must be followed by a company name (LLC, Inc, Corp, Ltd, etc.)
  - If no label found: Vendor field will be empty (intentional)

Invoice Number:
  - Looks for: "Invoice No", "Invoice Number", "Invoice #", "Invoice ID"
  - Also recognizes "# 12345" format

Date:
  - Looks for: "Invoice Date", "Date:", "Dated:", "Issue Date"
  - Supports multiple formats (DD/MM/YYYY, MM/DD/YYYY, etc.)

================================================================================
                         TROUBLESHOOTING
================================================================================

Problem: App won't start
Solution: Try running as Administrator, or check antivirus

Problem: All files show as "Scanned PDF"
Solution: Your PDFs are image-based. Ask vendor for digital invoices.

Problem: Missing fields in results
Solution: The PDF format may differ from common layouts.

Problem: Vendor field is always empty
Solution: Tool only extracts vendor when explicitly labeled.
          This is intentional to avoid extracting incorrect data.

Problem: Duplicate filenames
Solution: Tool automatically appends _1, _2, etc. to duplicates.

Problem: Files not being copied
Solution: 
  1. Make sure "Dry Run" is UNCHECKED
  2. Check you have write permission to output folder
  3. Check errors.log for specific issues

================================================================================
                        ANTIVIRUS WARNING
================================================================================

Some antivirus software may flag this application as unknown or suspicious.
This is a common FALSE POSITIVE for packaged Python applications.

If your antivirus blocks the app:
1. Check your antivirus quarantine/blocked items
2. Restore and allow the application
3. Add an exception if needed

The application is safe - it only reads PDFs and writes Excel/CSV files.
Your original PDF files are NEVER modified.

================================================================================
                    FOR SCANNED INVOICES
================================================================================

If you have scanned invoices (image-based PDFs), you'll need OCR first.
Options:
* Rescan using "Searchable PDF" or "OCR" mode on your scanner
* Use Adobe Acrobat's "Recognize Text" feature
* Use online PDF OCR services (Google, Adobe online)

After OCR processing, try InvFrog again.

HOW TO CHECK IF YOUR PDF IS MACHINE-GENERATED
---------------------------------------------
1. Open your PDF in any PDF viewer
2. Try to select/highlight text with your mouse
3. If you can copy-paste text, it should work
4. If you cannot select text, it's likely a scanned image

================================================================================
                         Version 2.0 | 2026
================================================================================
