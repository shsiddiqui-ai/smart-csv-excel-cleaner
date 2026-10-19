# Smart CSV & Excel Data Cleaner (Portfolio Project 1)

A Python starter tool that reads CSV/XLSX contact data, normalizes whitespace and email case, removes **exact duplicates after normalization**, checks email/phone **formats only**, and exports a clean CSV/XLSX file.

## Windows PowerShell setup

```powershell
cd D:\Python-Projects
mkdir smart-csv-excel-cleaner
# Extract the downloaded ZIP contents into D:\Python-Projects\smart-csv-excel-cleaner
cd D:\Python-Projects\smart-csv-excel-cleaner
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe cleaner.py
```

The last command creates `output/cleaned_contacts.xlsx` and shows summary counts.

To run your own file (with the four expected columns Name, Email, Phone, City):

```powershell
.\.venv\Scripts\python.exe cleaner.py --input "input/myfile.csv" --output "output/myfile_cleaned.xlsx"
```

## Limitations of v1

- No external identity, email deliverability, phone activity, Truecaller, or WhatsApp checks.
- Duplicate rule is **exact match on Name + Email + Phone + City after normalization**. Fuzzy duplicates are not yet supported.
- Does not modify the source file.
- Sample contacts are fictional.

## Planned extensions

1. Multiple deduplication rules (email, phone, or fuzzy review).
2. Invalid-row review report, Excel styling, and automated tests.
3. Flask drag-and-drop upload/download interface.
4. README screenshots, short recorded demo, GitHub portfolio page.
