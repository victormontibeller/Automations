"""Process-level exclusion for sending and delivery history maintenance."""

from __future__ import annotations

from contextlib import contextmanager
import fcntl
from pathlib import Path

from .errors import ReportError


@contextmanager
def run_lock(state_dir: Path):
    state_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
    with (state_dir / "send.lock").open("a") as handle:
        try:
            fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as exc:
            raise ReportError("Já existe uma execução de envio ou manutenção do histórico em andamento.") from exc
        try:
            yield
        finally:
            fcntl.flock(handle, fcntl.LOCK_UN)
