"""SQLite delivery history, interruption recovery and manual reconciliation."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import sqlite3

from .config import Settings
from .errors import ReportError
from .locking import run_lock
from .models import Report


class Ledger:
    def __init__(self, path: Path):
        self.db = sqlite3.connect(path)
        self.db.row_factory = sqlite3.Row
        self.db.execute('''CREATE TABLE IF NOT EXISTS deliveries (
            month TEXT NOT NULL, card TEXT NOT NULL, sheet TEXT NOT NULL,
            status TEXT NOT NULL CHECK(status IN ('sending','sent','failed','unknown')),
            recipient TEXT NOT NULL, message_id TEXT NOT NULL, source_hash TEXT NOT NULL,
            updated_at TEXT NOT NULL, note TEXT NOT NULL DEFAULT '',
            PRIMARY KEY (month,card,sheet))''')
        # Caller holds the process lock: no live sender can own these entries.
        self.db.execute("UPDATE deliveries SET status='unknown', note='Execução interrompida; conferir no Gmail.' WHERE status='sending'")
        self.db.commit()

    def close(self):
        self.db.close()

    def get(self, report: Report):
        return self.db.execute("SELECT * FROM deliveries WHERE month=? AND card=? AND sheet=?", (report.month, report.card, report.recipient.sheet)).fetchone()

    def claim(self, report: Report, message_id: str, source_hash: str):
        self.db.execute('''INSERT INTO deliveries VALUES (?,?,?,'sending',?,?,?,?, '')
            ON CONFLICT(month,card,sheet) DO UPDATE SET status='sending',recipient=excluded.recipient,
            message_id=excluded.message_id,source_hash=excluded.source_hash,updated_at=excluded.updated_at,note='' ''',
            (report.month, report.card, report.recipient.sheet, report.recipient.email, message_id, source_hash, datetime.now(timezone.utc).isoformat()))
        self.db.commit()

    def finish(self, report: Report, status: str, note: str = ""):
        self.db.execute("UPDATE deliveries SET status=?,note=?,updated_at=? WHERE month=? AND card=? AND sheet=?",
                        (status, note, datetime.now(timezone.utc).isoformat(), report.month, report.card, report.recipient.sheet))
        self.db.commit()


def maintain_ledger(settings: Settings, month: str | None, card: str | None, recipient: str | None, resolve: str | None) -> None:
    if not (settings.state_dir / "deliveries.sqlite3").exists():
        raise ReportError("Ainda não existe histórico de envios.")
    with run_lock(settings.state_dir):
        ledger = Ledger(settings.state_dir / "deliveries.sqlite3")
        try:
            if resolve:
                if not all((month, card, recipient)):
                    raise ReportError("--resolve exige --month, --card e --recipient.")
                rows = ledger.db.execute("SELECT * FROM deliveries WHERE month=? AND card=? AND sheet=?", (month, card, recipient)).fetchall()
                if len(rows) != 1 or rows[0]["status"] != "unknown":
                    raise ReportError("Só é possível resolver um envio com status unknown.")
                new_status = "sent" if resolve == "sent" else "failed"
                ledger.db.execute("UPDATE deliveries SET status=?,note=?,updated_at=? WHERE month=? AND card=? AND sheet=?",
                                  (new_status, "Conferido manualmente pelo proprietário.", datetime.now(timezone.utc).isoformat(), month, card, recipient))
                ledger.db.commit()
            rows = ledger.db.execute("SELECT month,card,sheet,status,recipient,updated_at,message_id FROM deliveries ORDER BY month,card,sheet").fetchall()
            print("mês | cartão | aba | estado | destinatário | atualização (UTC) | Message-ID")
            for row in rows:
                if (month is None or row["month"] == month) and (card is None or row["card"] == card) and (recipient is None or row["sheet"] == recipient):
                    print(" | ".join(str(value) for value in row))
        finally:
            ledger.close()
