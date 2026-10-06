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
    STATUS_MISSING,
    STATUS_RECONCILED,
    analyze,
    parse_date,
)
from automations.cupons_hospedes_config import (
    TRANSACTION_CODES_BY_HOTEL,
    transaction_codes_by_hotel,
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
    def test_internal_mapping_matches_the_homologated_workbook(self):
        self.assertEqual(
            {hotel: len(codes) for hotel, codes in TRANSACTION_CODES_BY_HOTEL.items()},
            {"TAIBA": 22, "CHARME": 19, "MAGNA": 6, "CUMBUCO": 27},
        )
        self.assertIn("2111", TRANSACTION_CODES_BY_HOTEL["TAIBA"])
        self.assertIn("2028", TRANSACTION_CODES_BY_HOTEL["CHARME"])
        self.assertIn("2002", TRANSACTION_CODES_BY_HOTEL["MAGNA"])
        self.assertIn("2051", TRANSACTION_CODES_BY_HOTEL["CUMBUCO"])

    def test_internal_mapping_returns_an_isolated_copy(self):
        mappings = transaction_codes_by_hotel()
        mappings["CHARME"].add("9999")

        self.assertNotIn("9999", TRANSACTION_CODES_BY_HOTEL["CHARME"])

    def test_matches_magna_account_using_check_prefix(self):
        accounts = {"10008370", "10008371"}

        self.assertEqual(
            _match_account("0018370", accounts, "MAGNA PRAIA"), "10008370"
        )

    def test_matches_magna_frigobar_account_using_check_prefix(self):
        accounts = {"60011071", "60011072"}

        self.assertEqual(
            _match_account("0061071", accounts, "MAGNA PRAIA"), "60011071"
        )

    def test_short_account_does_not_capture_larger_check_by_suffix(self):
        accounts = {"30", "10009330"}

        self.assertEqual(
            _match_account("0019330", accounts, "MAGNA PRAIA"), "10009330"
        )
        self.assertEqual(_match_account("30", accounts, "MAGNA PRAIA"), "30")

    def test_matches_charme_accounts_using_outlet_prefix(self):
        accounts = {
            "10016419",
            "20006868",
            "40003550",
            "60004631",
        }

        expected = {
            "0016419": "10016419",
            "0026868": "20006868",
            "0043550": "40003550",
            "0064631": "60004631",
        }
        for check, account in expected.items():
            with self.subTest(check=check):
                self.assertEqual(
                    _match_account(check, accounts, "CHARME HOSPEDAGEM"), account
                )

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
        paths = [Path("pdv.xlsx"), Path("journal.xlsx")]
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
                side_effect=("pdv", "journal"),
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
                "automations.conciliacao_cupons_hospedes.transaction_codes_by_hotel",
                return_value=mappings,
            ),
        ):
            result = analyze(paths)

        self.assertEqual(result.company, "CHARME HOSPEDAGEM")
        self.assertEqual(result.mapping, "CHARME")

    def test_missing_coupon_status_identifies_the_missing_source(self):
        paths = [Path("pdv.xlsx"), Path("journal.xlsx")]
        matched = _Coupon(
            "CHARME HOSPEDAGEM",
            "PDV",
            date(2026, 8, 1),
            "1",
            "440003175",
            "0100",
            "Hóspede 1",
            "Cupom",
            Decimal("13.00"),
        )
        missing = _Coupon(
            "CHARME HOSPEDAGEM",
            "PDV",
            date(2026, 8, 1),
            "2",
            "440003176",
            "0101",
            "Hóspede 2",
            "Cupom",
            Decimal("20.00"),
        )
        coupons = {
            (item.company, item.account, item.issue_date, item.document): item
            for item in (matched, missing)
        }
        journal = [
            _JournalRow(
                "2028", "440003175", date(2026, 8, 1), Decimal("13.00"), "0100"
            )
        ]

        with (
            patch(
                "automations.conciliacao_cupons_hospedes.identify_file",
                side_effect=("pdv", "journal"),
            ),
            patch(
                "automations.conciliacao_cupons_hospedes._read_pdv",
                return_value=coupons,
            ),
            patch(
                "automations.conciliacao_cupons_hospedes._read_journal",
                return_value=journal,
            ),
            patch(
                "automations.conciliacao_cupons_hospedes.transaction_codes_by_hotel",
                return_value={"CHARME": {"2028"}},
            ),
        ):
            result = analyze(paths)

        missing_result = next(item for item in result.coupons if item.document == "2")
        self.assertEqual(missing_result.status, STATUS_MISSING)
        self.assertIn("não localizado no Journal", missing_result.status)

    def test_reconciles_multiple_coupons_by_account_and_date_total(self):
        paths = [Path("pdv.xlsx"), Path("journal.xlsx")]
        coupons = [
            _Coupon(
                "CHARME HOSPEDAGEM",
                "Frigobar",
                date(2026, 9, 28),
                document,
                "60005005",
                "0304",
                "Hóspede",
                "Cupom",
                Decimal("57.00"),
            )
            for document in ("22585", "22586")
        ]
        journal = [
            _JournalRow(
                "2001", "0065005", date(2026, 9, 28), Decimal("57.00"), "0304"
            )
            for _ in range(2)
        ]

        with (
            patch(
                "automations.conciliacao_cupons_hospedes.identify_file",
                side_effect=("pdv", "journal"),
            ),
            patch(
                "automations.conciliacao_cupons_hospedes._read_pdv",
                return_value={
                    (item.company, item.account, item.issue_date, item.document): item
                    for item in coupons
                },
            ),
            patch(
                "automations.conciliacao_cupons_hospedes._read_journal",
                return_value=journal,
            ),
            patch(
                "automations.conciliacao_cupons_hospedes.transaction_codes_by_hotel",
                return_value={"CHARME": {"2001"}},
            ),
        ):
            result = analyze(paths)

        self.assertTrue(all(item.status == STATUS_RECONCILED for item in result.coupons))
        self.assertTrue(
            all(item.journal_value == Decimal("57.00") for item in result.coupons)
        )

    def test_reconciles_positive_reposting_after_reversal(self):
        paths = [Path("pdv.xlsx"), Path("journal.xlsx")]
        coupon = _Coupon(
            "CHARME HOSPEDAGEM",
            "Restaurante",
            date(2026, 9, 2),
            "20243",
            "10016493",
            "0309",
            "Hóspede",
            "Cupom",
            Decimal("6195.00"),
        )
        journal = [
            _JournalRow("2004", "0016493", date(2026, 9, 2), value, "0309")
            for value in (Decimal("-11065.00"), Decimal("6195.00"))
        ]

        with (
            patch(
                "automations.conciliacao_cupons_hospedes.identify_file",
                side_effect=("pdv", "journal"),
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
                "automations.conciliacao_cupons_hospedes.transaction_codes_by_hotel",
                return_value={"CHARME": {"2004"}},
            ),
        ):
            result = analyze(paths)

        self.assertEqual(result.coupons[0].status, STATUS_RECONCILED)
        self.assertEqual(result.coupons[0].journal_value, Decimal("6195.00"))

    def test_reconciles_one_exact_occurrence_among_repeated_postings(self):
        paths = [Path("pdv.xlsx"), Path("journal.xlsx")]
        coupon = _Coupon(
            "CARMEL TAÍBA",
            "Restaurante Cipó",
            date(2026, 9, 1),
            "25809",
            "80028710",
            "",
            "Hóspede",
            "Cupom",
            Decimal("63.00"),
        )
        journal = [
            _JournalRow("2111", "0028710", date(2026, 9, 1), value, "")
            for value in (
                Decimal("37.00"),
                Decimal("26.00"),
                Decimal("37.00"),
                Decimal("26.00"),
                Decimal("-37.00"),
                Decimal("-26.00"),
                Decimal("37.00"),
                Decimal("26.00"),
            )
        ]

        with (
            patch(
                "automations.conciliacao_cupons_hospedes.identify_file",
                side_effect=("pdv", "journal"),
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
                "automations.conciliacao_cupons_hospedes.transaction_codes_by_hotel",
                return_value={"TAIBA": {"2111"}},
            ),
        ):
            result = analyze(paths)

        self.assertEqual(result.coupons[0].status, STATUS_RECONCILED)
        self.assertEqual(result.coupons[0].journal_value, Decimal("63.00"))
