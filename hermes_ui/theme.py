from __future__ import annotations

import flet as ft

THEME_SEED = "#4C3427"
CARMEL_BROWN = "#4C3427"
CARMEL_TAUPE = "#9F9385"
CARMEL_STONE = "#DBD4CA"
CARMEL_IVORY = "#E8E7E1"
CARMEL_SAND = "#E7D0A7"
CARMEL_BLACK = "#1D1D1B"
CARMEL_TERRACOTTA = "#B14D1D"
CARMEL_GOLD = "#AE894C"
FONT_BODY = "Plus Jakarta Sans"
FONT_DISPLAY = "Cormorant"

BLUE = ft.Colors.PRIMARY
SURFACE = ft.Colors.SURFACE_CONTAINER_LOW
CARD = ft.Colors.SURFACE_CONTAINER
BORDER = ft.Colors.OUTLINE_VARIANT
TEXT = ft.Colors.ON_SURFACE
MUTED = ft.Colors.ON_SURFACE_VARIANT
GREEN = ft.Colors.GREEN_600
YELLOW = CARMEL_GOLD
RED = ft.Colors.ERROR
PAGE_BACKGROUND = ft.Colors.SURFACE
SIDEBAR_BACKGROUND = ft.Colors.SURFACE_CONTAINER_LOW
FOOTER_BACKGROUND = ft.Colors.SURFACE_CONTAINER
PROGRESS_BACKGROUND = ft.Colors.SURFACE_CONTAINER_HIGHEST
SELECTED_BACKGROUND = ft.Colors.PRIMARY_CONTAINER
SELECTED_FOREGROUND = ft.Colors.ON_PRIMARY_CONTAINER


LIGHT_SCHEME = ft.ColorScheme(
    primary=CARMEL_BROWN,
    on_primary="#FFFFFF",
    primary_container=CARMEL_SAND,
    on_primary_container=CARMEL_BROWN,
    secondary=CARMEL_TAUPE,
    on_secondary=CARMEL_BLACK,
    secondary_container=CARMEL_STONE,
    on_secondary_container=CARMEL_BROWN,
    tertiary=CARMEL_TERRACOTTA,
    on_tertiary="#FFFFFF",
    surface="#F8F6F2",
    on_surface=CARMEL_BLACK,
    on_surface_variant="#665B54",
    outline="#81746B",
    outline_variant=CARMEL_STONE,
    surface_container_lowest="#FFFFFF",
    surface_container_low="#F3F0EB",
    surface_container="#ECE8E2",
    surface_container_high="#E5E0D8",
    surface_container_highest=CARMEL_STONE,
    error="#BA1A1A",
    on_error="#FFFFFF",
)

DARK_SCHEME = ft.ColorScheme(
    primary=CARMEL_SAND,
    on_primary=CARMEL_BLACK,
    primary_container=CARMEL_BROWN,
    on_primary_container="#FFF7ED",
    secondary=CARMEL_STONE,
    on_secondary=CARMEL_BLACK,
    secondary_container="#51483F",
    on_secondary_container="#F3EEE8",
    tertiary="#E48B5F",
    on_tertiary=CARMEL_BLACK,
    surface=CARMEL_BLACK,
    on_surface="#F3EFEA",
    on_surface_variant="#CFC6BC",
    outline=CARMEL_TAUPE,
    outline_variant="#51483F",
    surface_container_lowest="#171614",
    surface_container_low="#211F1C",
    surface_container="#292622",
    surface_container_high="#332F2A",
    surface_container_highest="#3C3731",
    error="#FFB4AB",
    on_error="#690005",
)


def next_theme_mode(current: ft.ThemeMode) -> ft.ThemeMode:
    """Alterna entre os dois modos visuais suportados pela aplicação."""
    return ft.ThemeMode.LIGHT if current == ft.ThemeMode.DARK else ft.ThemeMode.DARK


def configure_page_appearance(page: ft.Page) -> None:
    """Aplica a base visual compartilhada sem sobrescrever o tema escolhido."""
    page.bgcolor = PAGE_BACKGROUND
    page.padding = 0
    page.fonts = {
        FONT_BODY: "fonts/PlusJakartaSans.ttf",
        FONT_DISPLAY: "fonts/Cormorant-SemiBold.ttf",
    }
    page.theme = ft.Theme(
        color_scheme_seed=THEME_SEED,
        color_scheme=LIGHT_SCHEME,
        font_family=FONT_BODY,
        use_material3=True,
    )
    page.dark_theme = ft.Theme(
        color_scheme_seed=THEME_SEED,
        color_scheme=DARK_SCHEME,
        font_family=FONT_BODY,
        use_material3=True,
    )
