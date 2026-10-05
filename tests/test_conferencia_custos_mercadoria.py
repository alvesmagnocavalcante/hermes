from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from openpyxl import Workbook, load_workbook

from automations.conferencia_custos_mercadoria import (
    INVENTORY_CODES,
    analyze,
    entry_postings,
    export_excel,
    final_balances,
    identify,
    unmapped_analytic_accounts,
)


def create_workbook(path: Path, header: list[str], rows: list[list[object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(header)
    for row in rows:
        sheet.append(row)
    workbook.save(path)


class MerchandiseCostsTest(TestCase):
    def test_inventory_configuration_includes_all_expected_accounts(self):
        expected = {
            "Alimentos",
            "Vinhos & Champanhe",
            "Alcoolicos",
            "Não Alcoolicos",
            "Frigobar",
            "Mimos Hospedes",
            "Amenitees",
            "Material de Higiene e Limpeza",
            "Material de Cama, Mesa e Banho",
            "Material de Escritório/Informatica",
            "Decoracao",
            "Eletroeletronicos",
            "Suprimentos de uso do Hospedes",
            "Material de Copa e Cozinha",
            "Uniforme",
            "Material de Manutenção de Edifícios e Instalações",
            "Material de Manutenção de Maquinas e Equipamentos",
            "Material de Manutenção da Piscina",
            "Materal de Reposicao",
            "Equipamento de Protecao",
            "SPA",
        }

        self.assertEqual(set(INVENTORY_CODES), expected)
        self.assertEqual(INVENTORY_CODES["Material de Cama, Mesa e Banho"], ("0803",))

    def test_reports_unmapped_analytic_stock_accounts(self):
        data = (
            ("DescricaoConta", "TipoContaDescricao", "SaldoAtual"),
            [
                ("ESTOQUES", "Sintético", Decimal("100")),
                ("Alimentos", "Analítico", Decimal("80")),
                ("Conta nova", "Analítico", Decimal("20")),
            ],
        )

        self.assertEqual(unmapped_analytic_accounts(data), ["Conta nova"])

    def test_identifies_current_reports_by_content_with_generic_names(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            documents = root / "configuracao.xlsx"
            entries = root / "balancete 1.xlsx"
            inventory = root / "inventario.xlsx"
            stock = root / "balancete 2.xlsx"

            create_workbook(
                documents,
                ["DESCRICAOTDESEMB", "VALORLANÇADO"],
                [["Alimentos", Decimal("100")]],
            )
            create_workbook(
                entries,
                ["DescricaoConta", "Debito", "SaldoAtual", "Historico"],
                [
                    [
                        "Alimentos",
                        Decimal("100"),
                        Decimal("150"),
                        "Lançamento Nota Fiscal Eletrônica de Mercadoria (Terceiros)",
                    ],
                    ["Vinhos & Champanhe", 0, 0, ""],
                    ["Bebidas Alcoolicas", 0, 0, ""],
                    ["Bebidas Nao Alcoolicas", 0, 0, ""],
                    ["Frigobar", 0, 0, ""],
                ],
            )
            create_workbook(
                inventory,
                ["GrupoCodigo", "SaldoValor"],
                [["01", Decimal("150")]],
            )
            create_workbook(
                stock,
                ["DescricaoConta", "Debito", "SaldoAtual", "Historico"],
                [
                    ["ESTOQUES", 0, Decimal("150"), ""],
                    ["Alimentos", 0, Decimal("150"), ""],
                    ["Material de Copa e Cozinha", 0, 0, ""],
                ],
            )

            files = identify([documents, entries, inventory, stock])
            rows = analyze([documents, entries, inventory, stock])

        self.assertEqual(
            set(files), {"documents", "entry_ledger", "inventory", "stock_ledger"}
        )
        alimentos_entry = next(
            row
            for row in rows
            if row.analysis == "Entradas" and row.account == "Alimentos"
        )
        self.assertEqual(alimentos_entry.source, Decimal("100"))
        self.assertEqual(alimentos_entry.accounting, Decimal("100"))

    def test_entry_comparison_ignores_cost_and_stock_movements(self):
        data = (
            ("DescricaoConta", "Debito", "Historico"),
            [
                (
                    "Alimentos",
                    Decimal("100"),
                    "Lançamento Nota Fiscal Eletrônica de Mercadoria (Terceiros)",
                ),
                (
                    "Alimentos",
                    Decimal("70"),
                    "Saída por transferência - integração de custo",
                ),
                ("Alimentos", Decimal("5"), "Ajuste de saldo negativo"),
            ],
        )

        self.assertEqual(entry_postings(data), {"Alimentos": Decimal("100")})

    def test_reads_balance_from_row_below_account_without_movement(self):
        data = (
            ("DescricaoConta", "SaldoAtual"),
            [("Alimentos", None), (None, Decimal("125.50"))],
        )

        self.assertEqual(final_balances(data), {"Alimentos": Decimal("125.50")})

    def test_analyze_keeps_final_balance_for_account_without_movement(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            documents = root / "DOCUMENTOSLANCADOS.xlsx"
            entries = root / "RAZAOANALITICOESTOQUEAB.xlsx"
            inventory = root / "INVENTARIOFISICO.xlsx"
            stock = root / "RAZAOANALITICOESTOQUES.xlsx"

            create_workbook(documents, ["DESCRICAOTDESEMB", "VALORLANÇADO"], [])
            create_workbook(entries, ["DescricaoConta", "Debito", "Historico"], [])
            create_workbook(
                inventory,
                ["GrupoCodigo", "SaldoValor"],
                [["01", Decimal("125.50")]],
            )
            create_workbook(
                stock,
                ["DescricaoConta", "SaldoAtual"],
                [["Alimentos", None], [None, Decimal("125.50")]],
            )

            rows = analyze([documents, entries, inventory, stock])
            alimentos = next(
                row
                for row in rows
                if row.analysis == "Saldo final" and row.account == "Alimentos"
            )

            self.assertEqual(alimentos.accounting, Decimal("125.50"))
            self.assertEqual(alimentos.status, "Conciliado")

            result_path = root / "resultado.xlsx"
            export_excel(rows, result_path)
            result = load_workbook(result_path, read_only=True)
            self.assertEqual(result.sheetnames, ["Entradas", "Saldo final"])
            self.assertEqual(
                [cell.value for cell in result["Entradas"][1]],
                ["Conta", "CAP", "Contabilidade", "Diferença", "Status"],
            )
            self.assertEqual(
                [cell.value for cell in result["Saldo final"][1]],
                ["Conta", "Inventário", "Contabilidade", "Diferença", "Status"],
            )
            result.close()
