from datetime import date
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import patch

from openpyxl import Workbook

from automations.conciliacao_cupons_hospedes import (
    _Coupon,
    _JournalRow,
    _match_account,
    _read_journal,
    analyze,
    parse_date,
)


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


class MappingSelectionTest(TestCase):
    def test_matches_magna_account_using_check_prefix(self):
        accounts = {"10008370", "10008371"}

        self.assertEqual(
            _match_account("0018370", accounts, "MAGNA PRAIA"), "10008370"
        )
        self.assertIsNone(_match_account("0018370", accounts, "CHARME HOSPEDAGEM"))

    def test_rejects_ambiguous_transformed_magna_account(self):
        accounts = {"10008370", "19998370"}

        self.assertIsNone(_match_account("0018370", accounts, "MAGNA PRAIA"))

    def test_matches_taiba_accounts_using_outlet_prefix(self):
        accounts = {
            "20008487",
            "30002960",
            "60006811",
            "70002991",
            "80026015",
        }

        expected = {
            "0048487": "20008487",
            "0032960": "30002960",
            "0066811": "60006811",
            "0012991": "70002991",
            "0026015": "80026015",
        }
        for check, account in expected.items():
            with self.subTest(check=check):
                self.assertEqual(_match_account(check, accounts, "CARMEL TAÍBA"), account)

    def test_prioritizes_mapping_named_for_identified_company(self):
        paths = [Path("pdv.xlsx"), Path("journal.xlsx"), Path("mapping.xlsx")]
        coupon = _Coupon(
            "CHARME HOSPEDAGEM",
            "PDV",
            date(2026, 8, 1),
            "1",
            "440003175",
            "0100",
            "Hóspede",
            "Cupom",
            Decimal("13.00"),
        )
        journal = [
            _JournalRow(
                "2028", "440003175", date(2026, 8, 1), Decimal("13.00"), "0100"
            )
        ]
        mappings = {
            "CHARME": {"2028"},
            "WIND": {"2028", "2050"},
        }

        with (
            patch(
                "automations.conciliacao_cupons_hospedes.identify_file",
                side_effect=("pdv", "journal", "mapping"),
            ),
            patch(
                "automations.conciliacao_cupons_hospedes._read_pdv",
                return_value={
                    (coupon.company, coupon.account, coupon.issue_date, coupon.document): coupon
                },
            ),
            patch(
                "automations.conciliacao_cupons_hospedes._read_journal",
                return_value=journal,
            ),
            patch(
                "automations.conciliacao_cupons_hospedes._read_mappings",
                return_value=mappings,
            ),
        ):
            result = analyze(paths)

        self.assertEqual(result.company, "CHARME HOSPEDAGEM")
        self.assertEqual(result.mapping, "CHARME")
