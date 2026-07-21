from __future__ import annotations

from abc import ABC, abstractmethod
from typing import TYPE_CHECKING

from automations.legacy_ui import ctk

if TYPE_CHECKING:
    from main import App


class Automation(ABC):
    """Contrato obrigatório para todas as automações do painel."""

    name: str
    PAGE_PADDING = 30
    TITLE_SIZE = 26
    TITLE_PADY = (22, 3)
    SUBTITLE_PADY = (0, 10)
    FILES_PADY = (10, 6)
    SECTION_PADY = (0, 8)
    CONSOLE_FONT_SIZE = 11
    PAGINATION_PADY = (7, 14)
    CHART_HEIGHT = 125

    def __init__(self, app: App, container: ctk.CTkFrame) -> None:
        self.app = app
        self.container = container

    @abstractmethod
    def render(self) -> None:
        """Desenha os controles da automação dentro de ``container``."""
