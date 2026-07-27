from decimal import Decimal
from pathlib import Path
from unittest import TestCase

from automations.conferencia_contas_receber import grouped, identify_ledger


class IdentifyLedgerTest(TestCase):
    header = ("DescricaoConta", "Debito", "Movimento", "Historico")

    def test_identifies_billing_ledger_by_content_regardless_of_filename(self):
        rows = [("Notas a Faturar", 10, 10, "Teste")]

        kind = identify_ledger(Path("relatorio qualquer.xlsx"), self.header, rows)

        self.assertEqual(kind, "razao_faturar")

    def test_identifies_commission_ledger_by_content_with_accents(self):
        rows = [("Comissão de Cartão de Crédito", 10, 10, "Teste")]

        kind = identify_ledger(Path("relatorio qualquer.xlsx"), self.header, rows)

        self.assertEqual(kind, "razao_comissao")

    def test_uses_normalized_filename_as_fallback(self):
        rows = [(None, 10, 10, "Teste")]

        billing = identify_ledger(
            Path("RAZÃO_notas-A-FATURAR.xlsx"), self.header, rows
        )
        commission = identify_ledger(
            Path("razao COMISSÃO cartão.xlsx"), self.header, rows
        )

        self.assertEqual(billing, "razao_faturar")
        self.assertEqual(commission, "razao_comissao")


class ClientGroupingTest(TestCase):
    header = ("Cliente", "Saldo")

    def test_consolidates_cvc_accounts_before_comparison(self):
        accounting = grouped(
            self.header,
            [
                ("CVC BRASIL", -107471.82),
                ("CVC BRASIL OPERADORA E AGENCIA", 12353.70),
                ("CVC BRASIL OPERADORA E AGENCIA DE VIAGENS", 146142.28),
            ],
            "Cliente",
            "Saldo",
        )
        financial = grouped(
            self.header,
            [("CVC Brasil Operadora e Agência de Viagens", 51024.16)],
            "Cliente",
            "Saldo",
        )

        self.assertEqual(accounting["CVC"], financial["CVC"])

    def test_does_not_merge_brt_and_bwt(self):
        result = grouped(
            self.header,
            [("BRT Operadora", 10), ("BWT Operadora", 20)],
            "Cliente",
            "Saldo",
        )

        self.assertEqual(result["BRT"][1], 10)
        self.assertEqual(result["BWTOPERADORA"][1], 20)

    def test_consolidates_related_trade_names(self):
        result = grouped(
            self.header,
            [
                ("DECOLAR.COM", 103700.37),
                ("DECOLAR.COM LTDA", 6499.08),
                ("Despegar-PAM", -3874.56),
            ],
            "Cliente",
            "Saldo",
        )

        self.assertEqual(result["DECOLARDESPEGAR"][1], Decimal("106324.89"))
