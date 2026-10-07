"""Individual MIME messages and guarded, ledger-backed batch delivery."""

from __future__ import annotations

from datetime import datetime, timezone
from email.message import EmailMessage
from email.utils import format_datetime, formataddr, make_msgid
from typing import Callable

from . import gmail
from .config import Settings, valid_email
from .errors import ReportError
from .gmail import GmailSendRejected
from .ledger import Ledger
from .locking import run_lock
from .models import Report
from .rendering import render_html, render_text, subject


def make_message(report: Report, sender: str, sender_name: str, target: str, *, cc: str | None = None, test: bool = False, message_id: str | None = None) -> EmailMessage:
    message = EmailMessage()
    message["From"] = formataddr((sender_name, sender))
    message["To"] = target
    if cc and cc.casefold() != target.casefold():
        message["Cc"] = cc
    message["Subject"] = ("[TESTE] " if test else "") + subject(report)
    message["Date"] = format_datetime(datetime.now(timezone.utc))
    message["Message-ID"] = message_id or make_msgid(domain=sender.split("@")[-1])
    message.set_content(render_text(report))
    message.add_alternative(render_html(report), subtype="html")
    return message


def send_reports(reports: list[Report], settings: Settings, source_hash: str, *, test_to: str | None = None, connector: Callable | None = None) -> tuple[int, int]:
    if test_to is not None and not valid_email(test_to):
        raise ReportError("Endereço de teste inválido.")
    missing = [r.recipient.sheet for r in reports if not valid_email(r.recipient.email)]
    if missing and test_to is None:
        raise ReportError("Preencha os e-mails antes de enviar: " + ", ".join(missing))
    if not reports:
        return 0, 0
    if not valid_email(settings.personal_copy_email):
        raise ReportError("Configure um personal_copy_email válido antes de enviar.")
    if any(not r.owner_name or r.owner_name == "[seu nome]" for r in reports):
        raise ReportError("Preencha owner_name na configuração antes de enviar: é o nome apresentado na saudação do Hermes.")
    with run_lock(settings.state_dir):
        ledger = None if test_to else Ledger(settings.state_dir / "deliveries.sqlite3")
        connection = None
        try:
            pending, skipped = [], 0
            for report in reports:
                prior = ledger.get(report) if ledger else None
                if prior and prior["status"] == "unknown":
                    raise ReportError(f"Envio de {report.recipient.sheet} tem resultado incerto. Confira o histórico e resolva antes de continuar.")
                if prior and prior["status"] == "sent":
                    skipped += 1
                else:
                    pending.append(report)
            if not pending:
                return 0, skipped
            connection = (connector or gmail.gmail_connect)(settings.google_token_file)
            sender = connection.email_address
            sent = 0
            for report in pending:
                message_id = make_msgid(domain=sender.split("@")[-1])
                message = make_message(report, sender, settings.sender_name, test_to or report.recipient.email,
                                       cc=settings.personal_copy_email,
                                       test=test_to is not None, message_id=message_id)
                if ledger:
                    ledger.claim(report, message_id, source_hash)
                try:
                    connection.send_message(message)
                except GmailSendRejected as exc:
                    if ledger:
                        ledger.finish(report, "failed", type(exc).__name__)
                    raise ReportError(f"Gmail API recusou o envio para {report.recipient.sheet}. Mensagens anteriores permanecem registradas.") from exc
                except Exception as exc:
                    if ledger:
                        ledger.finish(report, "unknown", type(exc).__name__)
                    raise ReportError(f"Resultado incerto ao enviar para {report.recipient.sheet}; confira o Gmail antes de repetir.") from exc
                if ledger:
                    ledger.finish(report, "sent")
                sent += 1
            return sent, skipped
        finally:
            try:
                if connection is not None:
                    try:
                        connection.close()
                    except Exception:
                        # Cleanup must not invalidate acceptance or invite a resend.
                        pass
            finally:
                if ledger:
                    ledger.close()
