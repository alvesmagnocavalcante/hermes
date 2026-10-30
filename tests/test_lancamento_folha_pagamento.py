from datetime import datetime
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from openpyxl import load_workbook
from reportlab.pdfgen import canvas

from automations.lancamento_folha_pagamento import (
    DEFAULT_TEMPLATE,
    Mappings,
    NON_ACCOUNTING_EVENTS,
    PayrollResult,
    PostingRow,
    UnparameterizedEvent,
    build_rates,
    extract_company,
    extract_liquid_total,
    extract_payslip_liquid,
    extract_payslip_period,
    find_unparameterized_events,
    identify_file,
    parse_monthly,
    parse_vacation_receipts,
    read_csv,
    read_mappings,
    required_sources,
    save_excel,
    select_parameterized_events,
)


class PayrollSourceSelectionTest(TestCase):
    def test_default_template_contains_additional_event_mappings(self):
        mappings = read_mappings(DEFAULT_TEMPLATE)

        self.assertEqual(mappings.events["271"], ("302020101", "201010101"))
        self.assertEqual(mappings.events["311"], ("101020104", "201010101"))
        self.assertEqual(mappings.events["359"], ("201010601", "201010103"))
        self.assertEqual(
            mappings.vacations["359"],
            ("Horas Férias Noturnas", "201010601", "201010103"),
        )
        self.assertEqual(
            mappings.vacations["392"],
            ("Med.Eve.Var.Abono Pecuniário", "201010601", "201010103"),
        )
        self.assertTrue({"359", "392"}.issubset(mappings.excluded_events))
        self.assertTrue(NON_ACCOUNTING_EVENTS.issubset(mappings.excluded_events))
        self.assertTrue(
            {"1904", "1914", "201", "206", "230"}.issubset(
                mappings.excluded_events
            )
        )
        self.assertIn("AJUDADETRANSPORTEESTAGIARIO", mappings.descriptions)

    def test_generates_additional_vacation_events(self):
        mappings = read_mappings(DEFAULT_TEMPLATE)
        employees = {"MARIA": ("101", "Governança", "0101 - Governança")}
        rows = [
            ["Funcionário", "", "", "Maria"],
            ["359", "", "", "", "125,50"],
            ["392", "", "", "", "300,00"],
        ]

        result = parse_vacation_receipts(rows, employees, mappings)

        self.assertEqual([row.event for row in result], ["359", "392"])
        self.assertEqual(
            [row.value for row in result],
            [Decimal("125.50"), Decimal("300.00")],
        )
        self.assertTrue(
            all(
                (row.debit, row.credit) == ("201010601", "201010103")
                for row in result
            )
        )

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
                "monthly_liquid",
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

    def test_accepts_eight_reports_when_there_are_vacations(self):
        grouped = {
            kind: [Path(f"{kind}.csv")]
            for kind in (
                "summary",
                "monthly_liquid",
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

    def test_identifies_both_liquid_reports_by_payroll_type(self):
        with TemporaryDirectory() as directory:
            base = Path(directory)
            advance = base / "arquivo_a.csv"
            monthly = base / "arquivo_b.csv"
            advance.write_text(
                "RELAÇÃO DE LÍQUIDOS A PAGAR\n"
                "Período: 01/10/2026 à 31/10/2026 - Tp: 8 - Adiantamento Salarial\n",
                encoding="utf-8",
            )
            monthly.write_text(
                "RELAÇÃO DE LÍQUIDOS A PAGAR\n"
                "Período: 01/10/2026 à 31/10/2026 - Tp: 1 - Mensal\n",
                encoding="utf-8",
            )

            self.assertEqual(identify_file(advance), "advance_liquid")
            self.assertEqual(identify_file(monthly), "monthly_liquid")

    def test_identifies_optional_wind_intermittent_liquid_by_filename(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "LIQUIDO DE INTERMITENTE WIND - 10.2026.csv"
            path.write_text(
                "RELAÇÃO DE LÍQUIDOS A PAGAR\n"
                "Período: 01/10/2026 à 31/10/2026 - Tp: 1 - Mensal\n"
                "Registros para Empresa 1;2.423,55\n",
                encoding="utf-8",
            )

            self.assertEqual(identify_file(path), "intermittent_liquid")
            self.assertEqual(
                extract_liquid_total(read_csv(path)), Decimal("2423.55")
            )

    def test_summary_with_monthly_type_is_not_mistaken_for_liquid_report(self):
        with TemporaryDirectory() as directory:
            summary = Path(directory) / "arquivo_generico.csv"
            summary.write_text(
                "RELAÇÃO DE CÁLCULO\n"
                "Período: 01/10/2026 à 31/10/2026 - Tp: 1 - Mensal\n",
                encoding="utf-8",
            )

            self.assertEqual(identify_file(summary), "summary")

    def test_extracts_company_total_from_liquid_report(self):
        rows = [
            ["Registros para Organograma 3", "3.751,29"],
            ["Registros para Empresa 92", "87.468,08"],
        ]

        self.assertEqual(extract_liquid_total(rows), Decimal("87468.08"))

    def test_identifies_and_extracts_payslip_pdf(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "contracheque.pdf"
            document = canvas.Canvas(str(path))
            document.drawString(40, 780, "Demonstrativo de Pagamento 10/2026")
            document.drawString(40, 740, "6.543,18 Total Liquido")
            document.save()

            self.assertEqual(identify_file(path), "payslip")
            self.assertEqual(extract_payslip_liquid(path), Decimal("6543.18"))
            self.assertEqual(extract_payslip_period(path), (2026, 10))

    def test_extracts_company_from_filial_header(self):
        rows = [
            ["0008-CM CENTRAL DE SERVICOS ADMINISTRATIVOS LTDA", "28/10/2026"],
            ["CNPJ/CPF: 35.428.047/0001-35", "Relação de Cálculo"],
            [
                "Filial: 1 CM CENTRAL DE SERVICOS ADMINISTRATIVOS LTDA "
                "- CNPJ/CPF: 35.428.047/0001-35"
            ],
        ]

        self.assertEqual(
            extract_company(rows),
            "CM CENTRAL DE SERVICOS ADMINISTRATIVOS LTDA",
        )

    def test_extracts_company_from_numbered_header_as_fallback(self):
        rows = [["0008-CM CENTRAL DE SERVICOS ADMINISTRATIVOS LTDA"]]

        self.assertEqual(
            extract_company(rows),
            "CM CENTRAL DE SERVICOS ADMINISTRATIVOS LTDA",
        )

    def test_monthly_event_without_mapping_is_reported(self):
        mappings = Mappings(
            descriptions=set(),
            events={},
            organograms={},
            organogram_codes={"001": ("0101", "Administração")},
            provisions={},
            vacations={},
            rate_totals={},
            excluded_events=frozenset(),
        )
        rows = [
            ["Organograma: 001 - Administração"],
            ["Situação: Normal"],
            ["Total IRRF"],
            ["999", "Evento novo", "", "150,25", ""],
        ]

        postings, _ = parse_monthly(rows, mappings)
        missing = find_unparameterized_events(postings)

        self.assertEqual(len(postings), 1)
        self.assertFalse(postings[0].ready)
        self.assertEqual(len(missing), 1)
        self.assertEqual(missing[0].event, "999")
        self.assertEqual(missing[0].value, Decimal("150.25"))
        self.assertEqual(
            find_unparameterized_events(postings, frozenset({"999"})), []
        )
        self.assertEqual(select_parameterized_events(postings, mappings), [])

    def test_unparameterized_event_does_not_reduce_result_quality(self):
        mappings = Mappings(
            descriptions={"SALARIO"},
            events={"1": ("D", "C")},
            organograms={},
            organogram_codes={},
            provisions={},
            vacations={},
            rate_totals={},
            excluded_events=frozenset(),
        )
        candidates = [
            PostingRow(
                "Folha mensal",
                10,
                "",
                "CC",
                "",
                "1",
                "Salário",
                "D",
                "C",
                Decimal("100"),
            ),
            PostingRow(
                "Folha mensal",
                11,
                "",
                "CC",
                "",
                "999",
                "Evento novo",
                "",
                "",
                Decimal("50"),
            ),
        ]

        result_rows = select_parameterized_events(candidates, mappings)

        self.assertEqual([row.event for row in result_rows], ["1"])
        self.assertTrue(all(row.ready for row in result_rows))
        self.assertEqual(
            [event.event for event in find_unparameterized_events(candidates)],
            ["999"],
        )

    def test_unknown_events_outside_target_section_do_not_change_exclusions(self):
        mappings = Mappings(
            descriptions={"SALARIO"},
            events={"1": ("D", "C")},
            organograms={},
            organogram_codes={},
            provisions={},
            vacations={},
            rate_totals={},
            excluded_events=frozenset(),
        )
        rows = [
            ["Organograma: 001 - Administração"],
            ["999", "Evento desconhecido", "", "50,00", ""],
            ["1", "Salário", "", "100,00", ""],
        ]

        postings, ignored = parse_monthly(rows, mappings)

        self.assertEqual(postings, [])
        self.assertEqual(ignored, 1)

    def test_excel_contains_liquid_and_event_validations(self):
        missing = UnparameterizedEvent(
            event="999",
            description="Evento novo",
            occurrences=1,
            value=Decimal("150.25"),
            source_lines=(4,),
        )
        result = PayrollResult(
            company="Empresa teste",
            period_end=datetime(2026, 10, 31),
            earnings=Decimal("300000"),
            deductions=Decimal("212531.92"),
            net_payable=Decimal("197842.61"),
            monthly_liquid=Decimal("191299.43"),
            payslip_liquid=Decimal("6543.18"),
            rows=[],
            ignored_rows=0,
            excluded_rows=0,
            files={
                "Líquido mensal": "mensal.csv",
                "Contracheque": "contracheque.pdf",
            },
            warnings=["Há 1 evento sem parametrização contábil neste mês: 999."],
            unparameterized_events=[missing],
        )
        with TemporaryDirectory() as directory:
            output = Path(directory) / "resultado.xlsx"
            save_excel(result, output)
            workbook = load_workbook(output, data_only=True, read_only=True)
            try:
                self.assertIn("Validacoes", workbook.sheetnames)
                sheet = workbook["Validacoes"]
                self.assertEqual(sheet["B4"].value, 197842.61)
                self.assertEqual(sheet["B7"].value, "Conciliado")
                self.assertEqual(sheet["A10"].value, "999")
            finally:
                workbook.close()
