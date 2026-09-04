from unittest import TestCase

from hermes_ui.presentation import normalized, status_kind


class PresentationStatusTest(TestCase):
    def test_normalizes_accents_and_case(self):
        self.assertEqual(normalized("CONCILIADO"), "conciliado")
        self.assertEqual(normalized("Fora do período"), "fora do periodo")

    def test_classifies_status_for_visual_feedback(self):
        self.assertEqual(status_kind("Conciliado"), "ok")
        self.assertEqual(status_kind("Cancelado"), "info")
        self.assertEqual(status_kind("Valor divergente"), "error")
