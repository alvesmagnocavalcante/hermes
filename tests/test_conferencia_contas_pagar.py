from decimal import Decimal
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch

from automations.conferencia_contas_pagar import entities, grouped, identify


class EntityGroupingTest(TestCase):
    header = ("Fornecedor", "Saldo")

    def test_consolidates_cvc_before_comparison(self):
        accounting = grouped(
            (
                self.header,
                [
                    ("CVC BRASIL", -107471.82),
                    ("CVC BRASIL OPERADORA E AGENCIA DE VIAGENS", 99947.12),
                ],
            ),
            "Fornecedor",
            "Saldo",
        )
        financial = grouped(
            (self.header, [("CVC Operadora", 7524.70)]),
            "Fornecedor",
            "Saldo",
        )

        result = entities("Fornecedores", accounting, financial)

        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].name, "CVC")
        self.assertEqual(result[0].accounting, Decimal("7524.70"))
        self.assertEqual(result[0].financial, Decimal("7524.7"))
        self.assertEqual(result[0].status, "Conciliado")

    def test_does_not_merge_brt_and_bwt(self):
        result = grouped(
            (self.header, [("BRT Operadora", 10), ("BWT Operadora", 20)]),
            "Fornecedor",
            "Saldo",
        )

        self.assertEqual(result["BRT"][1], Decimal("10"))
        self.assertEqual(result["BWTOPERADORA"][1], Decimal("20"))


class FileIdentificationTest(TestCase):
    def test_accepts_current_cap_report_names_using_headers(self):
        files = {
            "balancete subconta adto fornec.xlsx": ("DescricaoSubconta", "Saldo"),
            "balancete subconta fornec.xlsx": ("DescricaoSubconta", "Saldo"),
            "posicao por fornecedor.xlsx": (
                "Fornecedor",
                "Saldo",
                "DescContaContabil",
            ),
            "adiant fornec.xlsx": ("NomeFornecedor", "Saldo", "ValorTotalAdiantado"),
            "AGREGADO IRRF RETIDO.xlsx": (
                "IdAgregado",
                "Valor",
                "TratamentoFiscal",
            ),
            "AGREGADO CSRF RETIDO.xlsx": (
                "IdAgregado",
                "Valor",
                "TratamentoFiscal",
            ),
            "AGREGADO ISS RETIDO.xlsx": (
                "IdAgregado",
                "Valor",
                "TratamentoFiscal",
            ),
            "razao analitico.xlsx": ("DescricaoConta", "Movimento", "Historico"),
        }

        with patch(
            "automations.conferencia_contas_pagar.read",
            side_effect=lambda path: (files[path.name], []),
        ):
            result = identify([Path(name) for name in files])

        self.assertEqual(len(result), 8)

    def test_identifies_generic_reports_by_account_and_tax_content(self):
        common_tax_header = (
            "IdAgregado",
            "Valor",
            "TratamentoFiscal",
            "Descricao",
            "Historico",
        )
        files = {
            "balancete 1.xlsx": (
                ("DescricaoSubconta", "Saldo", "DescricaoConta"),
                [("Fornecedor A", 10, "Adiantamento a Fornecedores")],
            ),
            "balancete 2.xlsx": (
                ("DescricaoSubconta", "Saldo", "DescricaoConta"),
                [("Fornecedor B", 20, "Fornecedores de Mercadorias e Serviços")],
            ),
            "posicao.xlsx": (
                ("Fornecedor", "Saldo", "DescContaContabil"),
                [],
            ),
            "adiantamentos.xlsx": (
                ("NomeFornecedor", "Saldo", "ValorTotalAdiantado"),
                [],
            ),
            "alterador 1.xlsx": (
                common_tax_header,
                [(1, 10, "Lança Alterador de Decréscimo", "ISS Retido - CAP", "")],
            ),
            "alterador 2.xlsx": (
                common_tax_header,
                [
                    (
                        2,
                        20,
                        "Lança Alterador de Decréscimo",
                        "PIS/Cofins/CSLL Retido 5952 - CAP",
                        "",
                    )
                ],
            ),
            "alterador 3.xlsx": (
                common_tax_header,
                [(3, 30, "Lança Alterador de Decréscimo", "IRRF 1708 - CAP", "")],
            ),
            "razao.xlsx": (
                ("DescricaoConta", "Movimento", "Historico"),
                [],
            ),
        }

        with patch(
            "automations.conferencia_contas_pagar.read",
            side_effect=lambda path: files[path.name],
        ):
            result = identify([Path(name) for name in files])

        self.assertEqual(
            set(result),
            {
                "balancete_adto",
                "balancete_fornecedor",
                "posicao_fornecedor",
                "adiantamentos",
                "agregado_irrf",
                "agregado_csrf",
                "agregado_iss",
                "razao_impostos",
            },
        )
