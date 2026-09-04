from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from openpyxl import Workbook

from automations.conciliacao_cupons_hospedes import _read_journal, parse_date


class JournalDateParsingTest(TestCase):
    def test_accepts_date_with_slashes_and_two_digit_year(self):
        self.assertEqual(parse_date("01/08/26"), date(2026, 8, 1))

    def test_journal_keeps_rows_with_two_digit_year(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "journal.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(
                [
                    "TRX_CODE",
                    "REFERENCE",
                    "CASHIER_DEBIT",
                    "BUSINESS_FORMAT_DATE",
                    "ROOM",
                ]
            )
            sheet.append(
                [
                    2028,
                    "Room# 0100 : CHECK# 440003175 [1147]",
                    13,
                    "01/08/26",
                    "0100",
                ]
            )
            workbook.save(path)

            rows = _read_journal(path)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].posting_date, date(2026, 8, 1))
        self.assertEqual(rows[0].check, "440003175")
