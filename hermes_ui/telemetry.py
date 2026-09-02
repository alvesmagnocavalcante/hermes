from __future__ import annotations

import logging
import os
import re
from typing import Any

import logfire


LOGGER = logging.getLogger("hermes.telemetry")
SERVICE_NAME = "hermes"
SERVICE_VERSION = "0.1.0"
_configured = False

_SECRET_PATTERN = re.compile(
    r"(?i)\b(token|password|senha|secret|authorization)\b\s*[:=]\s*\S+"
)
_WINDOWS_PATH_PATTERN = re.compile(r"[A-Za-z]:\\(?:[^\\\s]+\\)*[^\s,;]+")
_LONG_IDENTIFIER_PATTERN = re.compile(r"\b\d{20,}\b")


def _enabled(name: str, default: bool = True) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().casefold() in {"1", "true", "yes", "on", "sim"}


def configure_telemetry() -> bool:
    """Configura o Logfire por credencial local ou token sem bloquear o HERMES."""
    global _configured
    _configured = False

    if not _enabled("HERMES_TELEMETRY_ENABLED"):
        LOGGER.info("Telemetria Logfire desabilitada por configuração.")
        return False

    try:
        logfire.configure(
            send_to_logfire="if-token-present",
            service_name=SERVICE_NAME,
            service_version=SERVICE_VERSION,
            environment=os.environ.get("LOGFIRE_ENVIRONMENT", "production"),
            console=False,
        )
        if _enabled("HERMES_LOGFIRE_SYSTEM_METRICS"):
            logfire.instrument_system_metrics()
    except Exception:
        _configured = False
        LOGGER.exception("Não foi possível configurar a telemetria Logfire.")
        return False

    _configured = True
    LOGGER.info(
        "Telemetria Logfire ativa: serviço=%s, ambiente=%s.",
        SERVICE_NAME,
        os.environ.get("LOGFIRE_ENVIRONMENT", "production"),
    )
    event(
        "info",
        "HERMES iniciado",
        telemetry="logfire",
        system_metrics=_enabled("HERMES_LOGFIRE_SYSTEM_METRICS"),
    )
    return True


def event(level: str, message: str, **attributes: Any) -> None:
    """Envia somente atributos operacionais previamente selecionados."""
    if not _configured:
        return
    try:
        logfire.log(level, message, attributes)
    except Exception:
        LOGGER.exception("Falha ao enviar evento para o Logfire.")


def safe_error_message(error: BaseException, limit: int = 500) -> str:
    """Resume uma falha operacional removendo caminhos, segredos e identificadores."""
    message = " ".join(str(error).split()) or "Erro sem mensagem detalhada."
    message = _SECRET_PATTERN.sub(r"\1=[omitido]", message)
    message = _WINDOWS_PATH_PATTERN.sub("[caminho omitido]", message)
    message = _LONG_IDENTIFIER_PATTERN.sub("[identificador omitido]", message)
    return message[:limit]


def authentication_event(result: str) -> None:
    labels = {
        "authorized": "Login autorizado",
        "rejected": "Login recusado",
        "blocked": "Login temporariamente bloqueado",
        "configuration_error": "Falha na configuração da autenticação",
    }
    event(
        "info" if result == "authorized" else "warn",
        "Autenticação HERMES: {authentication_result}",
        result=result,
        authentication_result=labels.get(result, result),
        event_category="authentication",
    )


def automation_event(
    *,
    automation: str,
    automation_name: str,
    hotel: str,
    success: bool,
    file_count: int,
    duration_seconds: float,
    execution_mode: str,
    total_bytes: int | None = None,
    record_count: int | None = None,
    reconciled_count: int | None = None,
    pending_count: int | None = None,
    informational_count: int | None = None,
    error_type: str | None = None,
    error: BaseException | None = None,
) -> None:
    result = "sucesso" if success else "falha"
    attributes: dict[str, Any] = {
        "event_category": "automation",
        "result": result,
        "automation_key": automation,
        "automation_name": automation_name,
        "hotel": hotel,
        "success": success,
        "file_count": file_count,
        "duration_seconds": round(duration_seconds, 3),
        "duration_ms": round(duration_seconds * 1000),
        "execution_mode": execution_mode,
    }
    if total_bytes is not None:
        attributes["total_bytes"] = total_bytes
    if record_count is not None:
        attributes["record_count"] = record_count
    if reconciled_count is not None:
        attributes["reconciled_count"] = reconciled_count
    if pending_count is not None:
        attributes["pending_count"] = pending_count
    if informational_count is not None:
        attributes["informational_count"] = informational_count
    if record_count and reconciled_count is not None:
        attributes["quality_percent"] = round(reconciled_count / record_count * 100, 1)
    if error_type is not None:
        attributes["error_type"] = error_type
    if error is not None:
        attributes["error_detail"] = safe_error_message(error)

    if success:
        message = (
            "Automação concluída: {automation_name} | {hotel} | "
            "{record_count} registros | {duration_seconds}s"
        )
    else:
        attributes.setdefault("error_type", "Erro")
        attributes.setdefault("error_detail", "Erro sem mensagem detalhada.")
        message = (
            "Automação falhou: {automation_name} | {hotel} | "
            "{error_type}: {error_detail}"
        )
    event("info" if success else "error", message, **attributes)


def export_event(
    *,
    automation: str,
    automation_name: str,
    hotel: str,
    output_format: str,
    success: bool,
    duration_seconds: float,
    error_type: str | None = None,
    error: BaseException | None = None,
) -> None:
    result = "sucesso" if success else "falha"
    attributes: dict[str, Any] = {
        "event_category": "export",
        "result": result,
        "automation_key": automation,
        "automation_name": automation_name,
        "hotel": hotel,
        "output_format": output_format,
        "success": success,
        "duration_seconds": round(duration_seconds, 3),
        "duration_ms": round(duration_seconds * 1000),
    }
    if error_type is not None:
        attributes["error_type"] = error_type
    if error is not None:
        attributes["error_detail"] = safe_error_message(error)

    if success:
        message = (
            "Exportação concluída: {automation_name} | {hotel} | "
            "{output_format} | {duration_seconds}s"
        )
    else:
        attributes.setdefault("error_type", "Erro")
        attributes.setdefault("error_detail", "Erro sem mensagem detalhada.")
        message = (
            "Exportação falhou: {automation_name} | {hotel} | "
            "{error_type}: {error_detail}"
        )
    event("info" if success else "error", message, **attributes)
