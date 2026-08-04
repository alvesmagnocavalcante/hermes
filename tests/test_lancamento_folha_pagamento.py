from datetime import datetime
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from automations.lancamento_folha_pagamento import (
    Mappings,
    PostingRow,
    build_rates,
    identify_file,
    required_sources,
)


class PayrollSourceSelectionTest(TestCase):
    def test_does_not_generate_health_and_dental_plan_rates(self):
        monthly = [
            PostingRow(
                "Folha mensal",
                1,
                "1 - Administração",
                "100",
                "Administração",
                "1",
                "SALÁRIO",
                "1",
                "2",
                Decimal("1000"),
            )
        ]
        mappings = Mappings(
            set(),
            {},
            {},
            {},
            {
                "REFPLANOODONTOLOGICO": ("1", "2"),
                "REFPLANODESAUDE": ("1", "2"),
                "REFINSSMENSAL": ("1", "2"),
                "REFFGTSMENSAL": ("1", "2"),
            },
            {},
            {
                "REF PLANO ODONTOLOGICO": Decimal("550.02"),
                "REF PLANO DE SAUDE": Decimal("8272.17"),
                "REF INSS MENSAL": Decimal("100"),
                "REF FGTS MENSAL": Decimal("80"),
            },
            frozenset(),
        )

        rows, _ = build_rates(
            monthly,
            mappings,
            Decimal("100"),
            Decimal("80"),
            datetime(2026, 7, 31),
        )

        descriptions = {row.description for row in rows}
        self.assertEqual(descriptions, {"REF INSS MENSAL", "REF FGTS MENSAL"})

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
