"""
Simple test script to verify parser logic without GUI.
"""
import os, sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import parser
from parser import NamingScheme


def test_filename_generation():
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
    # Note: date becomes 20231025 via parser logic
    assert name == "INV_INV-999_20231025.pdf"

    # Test 2: Vendor Scheme
    name = parser.generate_proposed_filename("scan001.pdf", mock_data, NamingScheme.VENDOR_NAME)
    print(f"Scheme: Vendor Name -> {name}")
    # Note: 'Acme Corp' has a space, which is allowed
    assert name == "Acme Corp_INV_INV-999_20231025.pdf"

    # Test 3: Missing Data
    mock_empty = {'invoice_number': None, 'date': None, 'vendor': None}
    name = parser.generate_proposed_filename("my_scan.pdf", mock_empty, NamingScheme.INVOICE_NUMBER)
    print(f"Scheme: Missing Data -> {name}")
    assert "unknown" in name

def test_path_calculation():
    # Mock Data
    mock_data = {
        'invoice_number': 'INV-999',
        'date': '20231025', # Parser ensures this format
        'vendor': 'Acme Corp'
    }

    path = parser.calculate_target_path("C:/Out", "file.pdf", mock_data, organize_by_month=True)
    # 20231025 -> 2023-10 folder
    assert "2023-10" in path

    path = parser.calculate_target_path("C:/Out", "file.pdf", mock_data, organize_by_month=False)
    assert "2023-10" not in path