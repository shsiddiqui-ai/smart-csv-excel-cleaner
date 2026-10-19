
# Smart CSV & Excel Data Cleaner

A Python and Flask application that automatically cleans contact data, removes duplicate records, validates basic contact formats, and generates professional Excel quality reports.

## Features

- Upload CSV and Excel (.xlsx) files
- Clean unnecessary spaces
- Standardize email capitalization
- Normalize phone number formatting
- Remove exact duplicate contact records
- Identify missing or incorrectly formatted emails
- Identify missing or incorrectly formatted phone numbers
- Generate a formatted Excel workbook
- Generate a separate CSV quality report
- Use a simple Flask web dashboard
- Basic protection against spreadsheet formula injection
- Automated unit tests

## Technology Stack

- Python
- Pandas
- OpenPyXL
- Flask
- unittest

## Example Results

Tested with 1,020 synthetic contact records:

| Metric | Result |
|---|---:|
| Input records | 1,020 |
| Duplicate records removed | 21 |
| Unique output records | 999 |
| Remaining exact duplicates | 0 |

The test dataset is synthetic and contains no real customer contacts.

## Getting Started

Clone the repository:

```bash
git clone https://github.com/shsiddiqui-ai/smart-csv-excel-cleaner.git
cd smart-csv-excel-cleaner
```

Create a virtual environment:

```bash
python -m venv .venv
```

Install dependencies:

```bash
pip install -r requirements.txt
```

## Web Dashboard

Start the Flask application:

```bash
python app.py
```

Open:

http://127.0.0.1:5001

Upload a CSV or XLSX contact file.

Required columns:

- Name
- Email
- Phone
- City

The dashboard downloads a cleaned Excel workbook with:

1. Cleaned Data
2. Quality Summary

Maximum upload size: 10 MB.

## Command-Line Usage

Run with the provided fictional sample:

```bash
python cleaner.py
```

Process your own file:

```bash
python cleaner.py --input input/sample_contacts.csv --output output/cleaned_contacts.xlsx
```

The command-line version creates a cleaned workbook and a separate CSV quality report.

## Automated Testing

Run:

```bash
python -m unittest tests.test_cleaner -v
```

Eight tests cover file creation, normalization, duplicate removal, format checks, missing columns, and report generation.

## Important Limitations

- Email validation checks syntax only. It does not verify deliverability.
- Phone validation checks basic syntax only. It does not verify active numbers.
- Duplicate detection uses exact matches across Name, Email, Phone and City after normalization.
- Fuzzy matching is not currently supported.
- This is a portfolio demonstration, not a production-ready public file-processing service.
- Spreadsheet safety controls are basic and require further security testing before public deployment.

## Planned Improvements

- Advanced duplicate detection
- Additional data-quality rules
- More upload security tests
- Processing statistics displayed on the web dashboard
- Dashboard screenshots and video demonstration

## Use Cases

Suitable as a starting point for:

- CRM contact-list cleanup
- Spreadsheet data standardization
- Duplicate record detection
- Contact data-quality reporting
- Python automation demonstrations

## Author

GitHub: https://github.com/shsiddiqui-ai
