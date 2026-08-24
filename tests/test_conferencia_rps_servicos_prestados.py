from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from automations.conferencia_rps_servicos_prestados import read_opera


class RPSServiceDescriptionsTest(TestCase):
    def test_accepts_additional_opera_services_by_normalized_description(self):
        xml = """<?xml version="1.0" encoding="utf-8"?>
<ROOT>
  <G_BILL_NO>
    <FOLIO_TYPE>NOTA</FOLIO_TYPE>
    <BILL_NO>123</BILL_NO>
    <BILL_GENERATION_DATE_CHAR>20/08/2026</BILL_GENERATION_DATE_CHAR>
    <ROOM>10</ROOM>
    <DISPLAY_NAME>Hóspede Teste</DISPLAY_NAME>
    <G_TRX_NO>
      <TRX_CODE>7777</TRX_CODE>
      <TRANSACTION_DESCRIPTION>Diária   Manual</TRANSACTION_DESCRIPTION>
      <FT_DEBIT>100,00</FT_DEBIT>
    </G_TRX_NO>
    <G_TRX_NO>
      <TRX_CODE>8888</TRX_CODE>
      <TRANSACTION_DESCRIPTION>SPA - Fleur Vigne Corporal</TRANSACTION_DESCRIPTION>
      <FT_DEBIT>50,00</FT_DEBIT>
    </G_TRX_NO>
    <G_TRX_NO>
      <TRX_CODE>9998</TRX_CODE>
      <TRANSACTION_DESCRIPTION>Produto não configurado</TRANSACTION_DESCRIPTION>
      <FT_DEBIT>25,00</FT_DEBIT>
    </G_TRX_NO>
  </G_BILL_NO>
</ROOT>
"""
        with TemporaryDirectory() as directory:
            path = Path(directory) / "opera.xml"
            path.write_text(xml, encoding="utf-8")

            rows = read_opera(path, set())

        self.assertEqual(rows["123"].value, Decimal("150.00"))
        self.assertIn("Diária Manual", rows["123"].services)
        self.assertIn("SPA - Fleur Vigne Corporal", rows["123"].services)
        self.assertNotIn("Produto não configurado", rows["123"].services)
