"""
Simple test script to verify parser logic without GUI.
"""
import os
import parser
from parser import NamingScheme

# Create dummy PDF for testing if none exists
dummy_pdf = "test_invoice.pdf"

print("--- Testing Filename Generation ---")

# Mock Data
mock_data = {
    'invoice_number': 'INV-999',
    'date': '2023-10-25',
    'vendor': 'Acme Corp',
    'total_amount': '150.00'
}

# Test 1: Default Scheme
name = parser.generate_proposed_filename("scan001.pdf", mock_data, NamingScheme.INVOICE_NUMBER)
print(f"Scheme: Invoice Num -> {name}")
assert name == "INV_INV-999_20231025.pdf"

# Test 2: Vendor Scheme
name = parser.generate_proposed_filename("scan001.pdf", mock_data, NamingScheme.VENDOR_NAME)
print(f"Scheme: Vendor Name -> {name}")
assert name == "Acme_Corp_INV_INV-999_20231025.pdf"

# Test 3: Missing Data
mock_empty = {'invoice_number': None, 'date': None, 'vendor': None}
name = parser.generate_proposed_filename("my_scan.pdf", mock_empty, NamingScheme.INVOICE_NUMBER)
print(f"Scheme: Missing Data -> {name}")
assert "unknown" in name

print("\n--- Testing Path Calculation ---")
path = parser.calculate_target_path("C:/Out", "file.pdf", mock_data, organize_by_month=True)
print(f"With Date Folders: {path}")
assert "2023" in path

path = parser.calculate_target_path("C:/Out", "file.pdf", mock_data, organize_by_month=False)
print(f"No Date Folders: {path}")
assert "2023" not in path

print("\nTests Completed Successfully.")