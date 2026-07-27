from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from openpyxl import Workbook

from automations.conferencia_cupons import read_file


class CouponReconciliationTest(TestCase):
    def test_aprovado_c_is_included_as_simphony(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "simphony.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Chave da NF", "Valor Total NF", "Status", "Data"])
            sheet.append(["CHAVE-1", 150, "Aprovado (C)", "01/07/2026"])
            sheet.append(["CHAVE-2", 80, "Cancelado", "01/07/2026"])
            workbook.save(path)

            source, values, cancelled = read_file(path)

            self.assertEqual(source, "Simphony")
            self.assertEqual(values["CHAVE-1"].value, Decimal("150"))
            self.assertNotIn("CHAVE-2", values)
            self.assertEqual(cancelled, 1)
