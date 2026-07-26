# syntax=docker/dockerfile:1

FROM ghcr.io/astral-sh/uv:0.11.21 AS uv

FROM python:3.12-slim AS builder

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy

WORKDIR /app

COPY --from=uv /uv /usr/local/bin/uv
COPY pyproject.toml uv.lock README.md ./

RUN uv sync --frozen --no-dev --no-install-project


FROM python:3.12-slim AS runtime

ENV PATH="/app/.venv/bin:$PATH" \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    FLET_FORCE_WEB_SERVER=true \
    FLET_SERVER_IP=0.0.0.0 \
    FLET_SERVER_PORT=8000 \
    FLET_MAX_UPLOAD_SIZE=104857600

RUN useradd --create-home --uid 10001 hermes

WORKDIR /app

COPY --from=builder /app/.venv /app/.venv
COPY --chown=hermes:hermes main.py ./
COPY --chown=hermes:hermes assets ./assets
COPY --chown=hermes:hermes automations ./automations
COPY --chown=hermes:hermes hermes_ui ./hermes_ui

USER hermes

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD ["python", "-c", "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/', timeout=3).read(1)"]

CMD ["python", "main.py"]
