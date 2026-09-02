from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from openpyxl import Workbook

from automations.conciliacao_receita_diarias import read_rules


class DailyRevenueRulesTest(TestCase):
    def test_includes_code_1011_only_for_taiba(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "codigos.xlsx"
            workbook = Workbook()
            workbook.remove(workbook.active)
            for hotel in ("Charme", "Cumbuco", "Magna", "Taiba"):
                sheet = workbook.create_sheet(hotel)
                sheet.append(["TRX_CODE", "D3", "DIÁRIA", "DIÁRIA MÉDIA"])
                sheet.append(["1000", "Diária", "SIM", "SIM"])
                sheet.append(["1011", "Código especial", None, None])
            workbook.save(path)

            for hotel in ("Cumbuco", "Magna", "Charme"):
                rules = read_rules(path, hotel)
                self.assertNotIn("1011", rules)
                self.assertIn("1000", rules)

            taiba = read_rules(path, "Taiba")
            self.assertIn("1011", taiba)
            self.assertTrue(taiba["1011"][1])
