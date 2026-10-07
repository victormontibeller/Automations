"""Immutable report data and exact Decimal balance calculations."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal

CENT = Decimal("0.01")
ZERO = Decimal("0.00")


@dataclass(frozen=True)
class Recipient:
    sheet: str
    participant: str
    email: str | None
    enabled: bool = True


@dataclass(frozen=True)
class Purchase:
    date: date
    description: str
    installment: str
    total: Decimal
    share: Decimal


@dataclass(frozen=True)
class Report:
    card: str
    month: str
    recipient: Recipient
    purchases: tuple[Purchase, ...]
    payments: tuple[Decimal, ...]
    owner_name: str | None = None
    payment_footer: str | None = None

    @property
    def spending(self) -> Decimal:
        return sum((p.share for p in self.purchases), ZERO)

    @property
    def balance(self) -> Decimal:
        return self.spending + sum(self.payments, ZERO)

    @property
    def has_movement(self) -> bool:
        return bool(self.purchases or self.payments)
