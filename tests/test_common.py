from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from openpyxl import Workbook

from automations.common import active_sheet_rows, parse_date


class CommonDateParsingTest(TestCase):
    def test_accepts_report_date_formats(self):
        cases = {
            "01/08/2026": date(2026, 8, 1),
            "01/08/26": date(2026, 8, 1),
            "01-08-2026": date(2026, 8, 1),
            "01-AUG-26": date(2026, 8, 1),
            "2026-08-01": date(2026, 8, 1),
        }
        for value, expected in cases.items():
            with self.subTest(value=value):
                self.assertEqual(parse_date(value), expected)

    def test_accepts_native_excel_dates(self):
        self.assertEqual(
            parse_date(datetime(2026, 8, 1, 10, 30)), date(2026, 8, 1)
        )
        self.assertEqual(parse_date(date(2026, 8, 1)), date(2026, 8, 1))

    def test_rejects_invalid_dates(self):
        self.assertIsNone(parse_date("não é uma data"))


class ActiveSheetRowsTest(TestCase):
    def test_returns_header_and_data_rows(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "fonte.xlsx"
            workbook = Workbook()
            workbook.active.append(["Código", "Valor"])
            workbook.active.append([1, 10])
            workbook.save(path)

            header, rows = active_sheet_rows(path)

        self.assertEqual(header, ("Código", "Valor"))
        self.assertEqual(rows, [(1, 10)])

    def test_rejects_empty_workbook_with_readable_error(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "vazia.xlsx"
            Workbook().save(path)

            with self.assertRaisesRegex(ValueError, "planilha vazia"):
                active_sheet_rows(path)
