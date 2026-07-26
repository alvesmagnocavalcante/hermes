from __future__ import annotations

import asyncio
import os
from collections.abc import Sequence
from pathlib import Path
from typing import Protocol


class UploadedFile(Protocol):
    name: str
    bytes: bytes | None
    path: str | None


def positive_int_env(name: str, default: int) -> int:
    raw = os.environ.get(name, str(default))
    try:
        value = int(raw)
    except ValueError as error:
        raise RuntimeError(f"{name} deve ser um número inteiro.") from error
    if value < 1:
        raise RuntimeError(f"{name} deve ser maior que zero.")
    return value


MAX_CONCURRENT_JOBS = positive_int_env("HERMES_MAX_CONCURRENT_JOBS", 2)
MAX_FILES_PER_JOB = positive_int_env("HERMES_MAX_FILES_PER_JOB", 20)
MAX_TOTAL_UPLOAD_SIZE = positive_int_env(
    "HERMES_MAX_TOTAL_UPLOAD_SIZE", 200 * 1024 * 1024
)
JOB_LIMITER = asyncio.Semaphore(MAX_CONCURRENT_JOBS)


def upload_size(files: Sequence[UploadedFile], web: bool) -> int:
    total = 0
    for selected in files:
        if web:
            if selected.bytes is None:
                raise ValueError(
                    f"O navegador não enviou o conteúdo de {selected.name}."
                )
            total += len(selected.bytes)
        else:
            if not selected.path:
                raise ValueError(
                    "Não foi possível acessar um ou mais arquivos selecionados."
                )
            total += Path(selected.path).stat().st_size
    return total


def validate_upload(files: Sequence[UploadedFile], web: bool) -> int:
    if len(files) > MAX_FILES_PER_JOB:
        raise ValueError(
            f"Selecione no máximo {MAX_FILES_PER_JOB} arquivos por operação."
        )
    total = upload_size(files, web)
    if total > MAX_TOTAL_UPLOAD_SIZE:
        limit_mib = MAX_TOTAL_UPLOAD_SIZE / (1024 * 1024)
        raise ValueError(
            f"O conjunto selecionado excede o limite total de {limit_mib:g} MiB."
        )
    return total
