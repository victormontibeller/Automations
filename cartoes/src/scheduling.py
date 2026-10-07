"""Card schedules and business month validation in the Brasilia timezone."""

from __future__ import annotations

from datetime import date, datetime, timedelta
import re
from zoneinfo import ZoneInfo

from .errors import ReportError

ZONE = ZoneInfo("America/Sao_Paulo")
CARDS = {"black": ("Black", 5), "latam": ("Latam", 20)}


def validate_month(month: str) -> str:
    if not re.fullmatch(r"[0-9]{4}-[0-9]{2}", month):
        raise ReportError("Mês deve estar no formato AAAA-MM.")
    try:
        date.fromisoformat(month + "-01")
    except ValueError as exc:
        raise ReportError("Mês inválido.") from exc
    return month


def previous_month(today: date) -> str:
    return (today.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")


def scheduled_month(card: str, now: datetime | None = None) -> str:
    now = now or datetime.now(ZONE)
    local = now.astimezone(ZONE)
    if local.day != CARDS[card][1]:
        raise ReportError("Execução fora do dia previsto. Para recuperar um envio, informe --month AAAA-MM manualmente.")
    if card == "latam":
        return local.strftime("%Y-%m")
    return previous_month(local.date())
