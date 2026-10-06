"""Parametrização interna dos TRX_CODE usados na conciliação de cupons."""

from __future__ import annotations


# Fonte homologada: planilha "DE PARA.xlsx" fornecida pela Contabilidade.
# A configuração fica centralizada para ser revisada sem alterar a conciliação.
TRANSACTION_CODES_BY_HOTEL: dict[str, frozenset[str]] = {
    "TAIBA": frozenset(
        {
            "2000", "2111", "2112", "2113", "2116", "2211", "2212", "2213",
            "2216", "2311", "2312", "2313", "2316", "2411", "2412", "2413",
            "2416", "3000", "3088", "4510", "4513", "4514",
        }
    ),
    "CHARME": frozenset(
        {
            "2000", "2001", "2003", "2004", "2005", "2006", "2007", "2028",
            "2029", "2030", "2031", "2040", "2041", "2042", "2043", "2048",
            "2049", "3088", "4000",
        }
    ),
    "MAGNA": frozenset({"2000", "2001", "2002", "2004", "2005", "2006"}),
    # A aba WIND da fonte homologada corresponde ao hotel Cumbuco.
    "CUMBUCO": frozenset(
        {
            "2000", "2001", "2002", "2003", "2004", "2005", "2006", "2007",
            "2008", "2016", "2018", "2019", "2021", "2028", "2029", "2030",
            "2031", "2040", "2041", "2042", "2043", "2048", "2049", "2050",
            "2051", "3088", "4000",
        }
    ),
}


def transaction_codes_by_hotel() -> dict[str, set[str]]:
    """Retorna uma cópia mutável para proteger a configuração global."""
    return {hotel: set(codes) for hotel, codes in TRANSACTION_CODES_BY_HOTEL.items()}
