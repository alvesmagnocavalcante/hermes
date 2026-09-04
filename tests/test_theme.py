from unittest import TestCase

import flet as ft

from hermes_ui.theme import (
    BORDER,
    CARD,
    PAGE_BACKGROUND,
    TEXT,
    next_theme_mode,
)


class ThemeConfigurationTest(TestCase):
    def test_switches_between_dark_and_light_modes(self):
        self.assertEqual(next_theme_mode(ft.ThemeMode.DARK), ft.ThemeMode.LIGHT)
        self.assertEqual(next_theme_mode(ft.ThemeMode.LIGHT), ft.ThemeMode.DARK)

    def test_main_palette_uses_adaptive_material_colors(self):
        self.assertEqual(PAGE_BACKGROUND, ft.Colors.SURFACE)
        self.assertEqual(CARD, ft.Colors.SURFACE_CONTAINER)
        self.assertEqual(BORDER, ft.Colors.OUTLINE_VARIANT)
        self.assertEqual(TEXT, ft.Colors.ON_SURFACE)
