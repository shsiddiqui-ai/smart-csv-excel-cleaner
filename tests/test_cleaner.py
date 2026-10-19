
"""Automated tests for Smart CSV & Excel Cleaner."""

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from cleaner import clean_file


class TestDataCleaner(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp_dir.name)

        self.input_file = self.folder / "contacts.csv"
        self.output_file = self.folder / "cleaned.xlsx"

        sample_data = pd.DataFrame([
            {
                "Name": " Ali Khan ",
                "Email": "ALI@EXAMPLE.COM",
                "Phone": "0300 1234567",
                "City": " Karachi "
            },
            {
                "Name": "Ali Khan",
                "Email": "ali@example.com",
                "Phone": "03001234567",
                "City": "Karachi"
            },
            {
                "Name": "Sara Ahmed",
                "Email": "wrong-email",
                "Phone": "123",
                "City": "Lahore"
            },
            {
                "Name": "Bilal Khan",
                "Email": "",
                "Phone": "",
                "City": "Multan"
            }
        ])

        sample_data.to_csv(self.input_file, index=False)

    def tearDown(self):
        self.temp_dir.cleanup()

    def run_cleaner(self):
        return clean_file(self.input_file, self.output_file)

    def test_excel_file_created(self):
        self.run_cleaner()
        self.assertTrue(self.output_file.exists())

    def test_duplicates_removed(self):
        result = self.run_cleaner()
        self.assertEqual(result["input_rows"], 4)
        self.assertEqual(result["duplicates_removed"], 1)
        self.assertEqual(result["output_rows"], 3)

    def test_email_normalization(self):
        self.run_cleaner()
        cleaned = pd.read_excel(
            self.output_file,
            sheet_name="Cleaned Data",
            dtype=str,
            keep_default_na=False
        )
        self.assertIn(
            "ali@example.com",
            cleaned["Email"].tolist()
        )

    def test_invalid_email_detected(self):
        self.run_cleaner()
        cleaned = pd.read_excel(
            self.output_file,
            sheet_name="Cleaned Data",
            dtype=str,
            keep_default_na=False
        )
        sara = cleaned.loc[
            cleaned["Name"] == "Sara Ahmed"
        ].iloc[0]

        self.assertEqual(
            sara["Email_Status"],
            "Invalid format"
        )

    def test_missing_phone_detected(self):
        self.run_cleaner()
        cleaned = pd.read_excel(
            self.output_file,
            sheet_name="Cleaned Data",
            dtype=str,
            keep_default_na=False
        )
        bilal = cleaned.loc[
            cleaned["Name"] == "Bilal Khan"
        ].iloc[0]

        self.assertEqual(bilal["Phone_Status"], "Missing")

    def test_quality_summary_exists(self):
        self.run_cleaner()
        with pd.ExcelFile(self.output_file) as workbook:
            self.assertEqual(
                workbook.sheet_names,
                ["Cleaned Data", "Quality Summary"]
            )

    def test_separate_quality_report_created(self):
        result = self.run_cleaner()
        report_path = Path(result["quality_report_path"])
        self.assertTrue(report_path.exists())

    def test_missing_required_column(self):
        pd.DataFrame([
            {"Name": "Ali", "Email": "ali@example.com"}
        ]).to_csv(self.input_file, index=False)

        with self.assertRaises(ValueError):
            self.run_cleaner()


if __name__ == "__main__":
    unittest.main()
