from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from automations.lancamento_folha_pagamento import (
    identify_file,
    required_sources,
)


class PayrollSourceSelectionTest(TestCase):
    def test_requires_irrf_when_there_are_no_vacation_reports(self):
        grouped = {
            kind: [Path(f"{kind}.csv")]
            for kind in (
                "summary",
                "inss",
                "fgts",
                "irrf",
                "vacation_provision",
                "thirteenth_provision",
            )
        }

        self.assertEqual(required_sources(grouped), set(grouped))

    def test_requires_both_vacation_reports_together(self):
        grouped = {"vacation_receipt": [Path("recibo.csv")]}

        with self.assertRaisesRegex(ValueError, "Líquido de férias"):
            required_sources(grouped)

    def test_accepts_seven_reports_when_there_are_vacations(self):
        grouped = {
            kind: [Path(f"{kind}.csv")]
            for kind in (
                "summary",
                "inss",
                "fgts",
                "vacation_receipt",
                "vacation_liquid",
                "vacation_provision",
                "thirteenth_provision",
            )
        }

        self.assertEqual(required_sources(grouped), set(grouped))

    def test_identifies_monthly_irrf_report(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "irrf.csv"
            path.write_text(
                "Relação de IRRF dos Empregados;Competência\n",
                encoding="utf-8",
            )

            self.assertEqual(identify_file(path), "irrf")
