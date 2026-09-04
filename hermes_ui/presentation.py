from __future__ import annotations

import unicodedata


def normalized(value: str) -> str:
    """Normaliza textos apresentados pela interface para classificação."""
    text = unicodedata.normalize("NFKD", value)
    return "".join(char for char in text if not unicodedata.combining(char)).casefold()


def status_kind(status: str) -> str:
    """Mapeia os status das automações para os três estados visuais da interface."""
    value = normalized(status)
    if any(
        word in value
        for word in (
            "conciliad",
            "no prazo",
            "em dia",
            "com movimento",
            "pronto",
            "processado",
            "cobrado",
            "dados completos",
        )
    ):
        return "ok"
    if any(
        word in value
        for word in (
            "cancelad",
            "fora do periodo",
            "sem movimento",
            "informativ",
            "alerta",
        )
    ):
        return "info"
    return "error"
