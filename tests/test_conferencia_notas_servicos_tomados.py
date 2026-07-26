from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from openpyxl import Workbook

from automations.conferencia_notas_servicos_tomados import analyze


class ServiceNotesTest(TestCase):
    def test_uses_iss_devido_and_ignores_date_and_provider_name_differences(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            cap = root / "cap (WIND).xlsx"
            tax = root / "iss (WIND).xlsx"
            city = root / "prefeitura (WIND).csv"

            workbook = Workbook()
            sheet = workbook.active
            sheet.append(
                [
                    "RazaoSocialFornecedor",
                    "DocumentoPrincipalFornecedor",
                    "Numero",
                    "DataEmissao",
                    "ValorBruto",
                    "StatusBPM",
                    "EmpresaNomeResumido",
                ]
            )
            sheet.append(
                [
                    "PRESTADOR NO CAP",
                    "12345678000190",
                    "10",
                    "01/07/2026",
                    100,
                    "BPM APROVADO",
                    "CARMEL CUMBUCO",
                ]
            )
            workbook.save(cap)

            workbook = Workbook()
            sheet = workbook.active
            sheet.append(
                [
                    "DocumentoPrincipalFornecedor",
                    "NumeroDocumento",
                    "ValorBaseCalculo",
                    "Valor",
                ]
            )
            sheet.append(["12345678000190", "10", 100, 5])
            workbook.save(tax)

            city.write_text(
                "Nº NFS-e;Data Hora NFE;CPF/CNPJ do Prestador;"
                "Razão Social do Prestador;Valor dos Serviços;"
                "ISS Retido;ISS devido\n"
                "10;02/07/2026;12345678000190;"
                "NOME DIFERENTE NA PREFEITURA;100,00;N;5,00\n",
                encoding="latin1",
            )

            result = analyze([cap, tax, city])
            row = result.rows[0]

            self.assertTrue(row.reconciled)
            self.assertEqual(result.expected_hotel, "CARMELCUMBUCO")
            self.assertEqual(row.iss, Decimal("5.00"))
            self.assertEqual(row.provider, "NOME DIFERENTE NA PREFEITURA")
            self.assertEqual(row.cap_provider, "PRESTADOR NO CAP")
            self.assertNotIn("Data divergente", row.status)
            self.assertNotIn("Razão social divergente", row.status)
