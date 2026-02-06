import parser
from parser import NamingScheme


def test_filename_generation_invoice_number():
    mock_data = {
        'invoice_number': 'INV-999',
        'date': '2023-10-25',
        'vendor': 'Acme Corp',
        'total_amount': '150.00'
    }

    name = parser.generate_proposed_filename(
        "scan001.pdf", mock_data, NamingScheme.INVOICE_NUMBER
    )

    assert name == "INV_INV-999_20231025.pdf"


def test_filename_generation_vendor():
    mock_data = {
        'invoice_number': 'INV-999',
        'date': '2023-10-25',
        'vendor': 'Acme Corp'
    }

    name = parser.generate_proposed_filename(
        "scan001.pdf", mock_data, NamingScheme.VENDOR_NAME
    )

    assert name.startswith("Acme_Corp")  # safer than exact match


def test_missing_data_fallback():
    mock_empty = {'invoice_number': None, 'date': None, 'vendor': None}

    name = parser.generate_proposed_filename(
        "my_scan.pdf", mock_empty, NamingScheme.INVOICE_NUMBER
    )

    assert "unknown" in name.lower()


def test_path_with_month():
    mock_data = {'date': '2023-10-25', 'vendor': 'Acme Corp'}

    path = parser.calculate_target_path(
        "/Out", "file.pdf", mock_data, organize_by_month=True
    )

    assert "2023-10" in path


def test_path_without_month():
    mock_data = {'date': '2023-10-25', 'vendor': 'Acme Corp'}

    path = parser.calculate_target_path(
        "/Out", "file.pdf", mock_data, organize_by_month=False
    )

    assert "2023-10" not in path
