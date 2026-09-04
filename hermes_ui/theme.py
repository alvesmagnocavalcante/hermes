from __future__ import annotations

import flet as ft

THEME_SEED = "#2383C4"
BLUE = ft.Colors.PRIMARY
SURFACE = ft.Colors.SURFACE_CONTAINER_LOW
CARD = ft.Colors.SURFACE_CONTAINER
BORDER = ft.Colors.OUTLINE_VARIANT
TEXT = ft.Colors.ON_SURFACE
MUTED = ft.Colors.ON_SURFACE_VARIANT
GREEN = ft.Colors.GREEN_600
YELLOW = ft.Colors.AMBER_600
RED = ft.Colors.ERROR
PAGE_BACKGROUND = ft.Colors.SURFACE
SIDEBAR_BACKGROUND = ft.Colors.SURFACE_CONTAINER_LOW
FOOTER_BACKGROUND = ft.Colors.SURFACE_CONTAINER
PROGRESS_BACKGROUND = ft.Colors.SURFACE_CONTAINER_HIGHEST
SELECTED_BACKGROUND = ft.Colors.PRIMARY_CONTAINER
SELECTED_FOREGROUND = ft.Colors.ON_PRIMARY_CONTAINER


def next_theme_mode(current: ft.ThemeMode) -> ft.ThemeMode:
    """Alterna entre os dois modos visuais suportados pela aplicação."""
    return ft.ThemeMode.LIGHT if current == ft.ThemeMode.DARK else ft.ThemeMode.DARK


def configure_page_appearance(page: ft.Page) -> None:
    """Aplica a base visual compartilhada sem sobrescrever o tema escolhido."""
    page.bgcolor = PAGE_BACKGROUND
    page.padding = 0
    page.theme = ft.Theme(color_scheme_seed=THEME_SEED, font_family="Segoe UI")
    page.dark_theme = ft.Theme(color_scheme_seed=THEME_SEED, font_family="Segoe UI")
