"""Executa uma verificação rápida das automações com exemplos completos."""

from __future__ import annotations

import sys
import tempfile
from importlib import import_module
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT_DIR))

SPECS = import_module("hermes_ui.registry").SPECS


EXAMPLE_DIRS = {
    "receita": "ATIVIDADE 1 - CONCILIACAO RECEITA",
    "folha": "ATIVIDADE 2 - FOLHA DE PAGAMENTO",
    "entrada": "ATIVIDADE 4 - CONFERENCIA NOTAS EM ATRASO",
    "cupons": "ATIVIDADE 5 - CUPONS",
    "rps": "ATIVIDADE 6 - PRESTADOS",
    "servicos": "ATIVIDADE 7 - TOMADOS",
    "receber": "ATIVIDADE 8 - CONFERENCIA DO CONTAS A RECEBER",
    "pagar": "ATIVIDADE 9 - CONFERENCIA DO CONTAS A PAGAR",
    "custos": "ATIVIDADE 10 - CONFERENCIA CUSTOS",
}

EXTENSIONS = {"Excel": "xlsx", "PDF": "pdf", "CSV": "csv"}


def validate() -> None:
    specs = {spec.key: spec for spec in SPECS}
    examples_root = ROOT_DIR / "PROCESSOS AUTOMAÇÃO"

    with tempfile.TemporaryDirectory(prefix="hermes_validation_") as temp_dir:
        output_dir = Path(temp_dir)
        for key, directory_name in EXAMPLE_DIRS.items():
            spec = specs[key]
            source_dir = examples_root / directory_name
            paths = sorted(path for path in source_dir.iterdir() if path.is_file())
            if not paths:
                raise FileNotFoundError(f"Sem arquivos de exemplo para {spec.name}.")

            result = spec.analyze(paths, "Magna")
            records = spec.rows(result)
            exports = []
            for output_format in spec.formats:
                output = output_dir / f"{key}.{EXTENSIONS[output_format]}"
                spec.export(result, output, output_format)
                if not output.is_file() or output.stat().st_size == 0:
                    raise RuntimeError(f"{spec.name} não gerou {output_format}.")
                exports.append(output_format)

            print(
                f"OK  {spec.name}: {len(records)} registro(s); "
                f"exportações {', '.join(exports)}"
            )


if __name__ == "__main__":
    validate()
