from __future__ import annotations

import hmac
import threading
import time
from collections import deque
from pathlib import Path
from typing import Callable

import bcrypt


class AuthenticationConfigurationError(RuntimeError):
    """Raised when the local credential file cannot be used."""


def _read_credentials(path: Path) -> list[tuple[str, bytes]]:
    if not path.is_file():
        raise AuthenticationConfigurationError(
            "Arquivo de credenciais não encontrado. Contate o administrador."
        )

    credentials: list[tuple[str, bytes]] = []
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as error:
        raise AuthenticationConfigurationError(
            "Não foi possível ler o arquivo de credenciais."
        ) from error

    for line in lines:
        if not line.strip() or line.lstrip().startswith("#") or ":" not in line:
            continue
        username, password_hash = line.split(":", 1)
        if username and password_hash:
            credentials.append((username, password_hash.encode("ascii")))

    if not credentials:
        raise AuthenticationConfigurationError(
            "O arquivo de credenciais não possui usuários válidos."
        )
    return credentials


def verify_credentials(path: Path, username: str, password: str) -> bool:
    """Validate a user against bcrypt entries from an Apache htpasswd file."""
    credentials = _read_credentials(path)
    selected_hash: bytes | None = None
    for stored_username, password_hash in credentials:
        if hmac.compare_digest(stored_username, username):
            selected_hash = password_hash
            break

    # Também executa bcrypt para usuário inexistente, reduzindo diferença de tempo.
    verification_hash = selected_hash or credentials[0][1]
    try:
        password_matches = bcrypt.checkpw(
            password.encode("utf-8"), verification_hash
        )
    except (TypeError, ValueError):
        raise AuthenticationConfigurationError(
            "O arquivo de credenciais contém um hash incompatível."
        ) from None
    return selected_hash is not None and password_matches


class LoginAttemptLimiter:
    """Process-local sliding-window limiter keyed by the connecting client."""

    def __init__(
        self,
        max_attempts: int = 5,
        window_seconds: int = 300,
        clock: Callable[[], float] = time.monotonic,
    ):
        self.max_attempts = max_attempts
        self.window_seconds = window_seconds
        self._clock = clock
        self._attempts: dict[str, deque[float]] = {}
        self._lock = threading.Lock()

    def _active_attempts(self, key: str, now: float) -> deque[float]:
        attempts = self._attempts.setdefault(key, deque())
        threshold = now - self.window_seconds
        while attempts and attempts[0] <= threshold:
            attempts.popleft()
        if not attempts:
            self._attempts.pop(key, None)
            return deque()
        return attempts

    def retry_after(self, key: str) -> int:
        with self._lock:
            now = self._clock()
            attempts = self._active_attempts(key, now)
            if len(attempts) < self.max_attempts:
                return 0
            return max(1, int(attempts[0] + self.window_seconds - now + 0.999))

    def record_failure(self, key: str) -> int:
        with self._lock:
            now = self._clock()
            attempts = self._active_attempts(key, now)
            if not attempts:
                attempts = self._attempts.setdefault(key, deque())
            attempts.append(now)
            if len(attempts) < self.max_attempts:
                return 0
            return max(1, int(attempts[0] + self.window_seconds - now + 0.999))

    def record_success(self, key: str) -> None:
        with self._lock:
            self._attempts.pop(key, None)


LOGIN_ATTEMPT_LIMITER = LoginAttemptLimiter()
