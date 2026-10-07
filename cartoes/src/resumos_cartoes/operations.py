"""Small, content-free operational results; not a delivery ledger."""
from dataclasses import dataclass
import io
import json
import math
import os
from typing import Literal
from uuid import UUID

from .errors import ReportError
from .scheduling import CARDS, validate_month

Mode = Literal["preview", "send", "test", "history", "resolve"]
Status = Literal["succeeded", "failed"]


def safe_card(value: object) -> str | None:
    return value if type(value) is str and value in CARDS else None


def safe_month(value: object) -> str | None:
    if type(value) is not str:
        return None
    try:
        return validate_month(value)
    except ReportError:
        return None


def _validate(event: str, fields: dict) -> None:
    """Reject entire malformed records, including future accidental extra fields."""
    common = {"run_id", "mode", "card", "month"}
    finished = {"status", "sent", "skipped", "no_movement", "previews",
                "elapsed_seconds", "error_category", "alert_status"}
    if type(event) is not str or event not in ("run_started", "run_finished"):
        raise ValueError("Invalid operational event")
    if fields.keys() != (common if event == "run_started" else common | finished):
        raise ValueError("Invalid operational fields")
    run_id = fields["run_id"]
    if type(run_id) is not str or str(UUID(run_id)) != run_id or UUID(run_id).version != 4:
        raise ValueError("Invalid run identifier")
    if type(fields["mode"]) is not str or fields["mode"] not in ("preview", "send", "test", "history", "resolve"):
        raise ValueError("Invalid mode")
    for name, sanitize in (("card", safe_card), ("month", safe_month)):
        if fields[name] is not None and (type(fields[name]) is not str or sanitize(fields[name]) is None):
            raise ValueError("Invalid period")
    if event == "run_finished":
        if type(fields["status"]) is not str or fields["status"] not in ("succeeded", "failed"):
            raise ValueError("Invalid status")
        if fields["error_category"] is not None and (type(fields["error_category"]) is not str or fields["error_category"] not in ("report_error", "operational_error")):
            raise ValueError("Invalid error category")
        if type(fields["alert_status"]) is not str or fields["alert_status"] not in ("not_requested", "accepted", "failed"):
            raise ValueError("Invalid alert status")
        for name in ("sent", "skipped", "no_movement", "previews"):
            value = fields[name]
            if value is not None and (type(value) is not int or value < 0):
                raise ValueError("Invalid count")
        elapsed = fields["elapsed_seconds"]
        if type(elapsed) not in (int, float) or not math.isfinite(elapsed) or elapsed < 0:
            raise ValueError("Invalid duration")


def emit_json(stream, event: str, **fields) -> None:
    """Best-effort, allowlisted diagnostics, never part of a send transaction."""
    try:
        _validate(event, fields)
        line = json.dumps({"event": event, **fields}, allow_nan=False) + "\n"
        if type(stream) is io.TextIOWrapper:
            try:
                fd = stream.fileno()
            except io.UnsupportedOperation:
                pass  # In-memory text wrappers still use their stream interface.
            else:
                # Native stderr must not retain failed JSON writes in its buffer:
                # a later interpreter-shutdown flush would replace exit 0 with 120.
                pending = line.encode(stream.encoding, stream.errors or "strict")
                while pending:
                    written = os.write(fd, pending)
                    if written <= 0:
                        return
                    pending = pending[written:]
                return
        stream.write(line)
        stream.flush()
    except Exception:
        # A broken logging sink must not change acceptance or trigger a resend.
        pass


@dataclass(frozen=True)
class RunResult:
    """One completed request; nullable counts mean unavailable/not applicable.

    A failed batch has unknown sent/skipped counts, not a claim of zero sends.
    Provider acceptance is not confirmation of delivery or reading.
    """
    run_id: str
    status: Status
    mode: Mode
    card: str | None
    month: str | None
    sent: int | None
    skipped: int | None
    no_movement: int | None
    previews: int | None
    elapsed_seconds: float
    error_category: str | None = None
    alert_status: str = "not_requested"

    @property
    def exit_code(self) -> int:
        return 0 if self.status == "succeeded" else 1
