"""Best-effort owner-only error notifications, without recursive retries."""

from __future__ import annotations

from datetime import datetime, timezone
from email.message import EmailMessage
from email.utils import format_datetime, formataddr, make_msgid

from . import gmail
from .config import DEFAULT_SENDER_NAME, Settings, hermes_home


def notify_owner(settings: Settings | None, card: str | None,
                 month: str | None, stage: str, reason: str) -> None:
    """One best-effort alert to the authenticated owner; never use recipients."""
    token_file = settings.google_token_file if settings else hermes_home() / "google_token.json"
    connection = gmail.gmail_connect(token_file)
    user = connection.email_address
    message = EmailMessage()
    display_name = settings.sender_name if settings else DEFAULT_SENDER_NAME
    message["From"] = formataddr((display_name, user))
    message["To"] = user
    copy_email = settings.personal_copy_email if settings else None
    if copy_email and copy_email.casefold() != user.casefold():
        message["Cc"] = copy_email
    message["Subject"] = "[ERRO] Resumos dos cartões" + (f" — {card.upper()}" if card else "") + (f" {month}" if month else "")
    message["Date"] = format_datetime(datetime.now(timezone.utc))
    message["Message-ID"] = make_msgid(domain=user.split("@")[-1])
    outcome = (
        "A execução foi interrompida durante a etapa de envio. Alguns resumos podem já ter sido aceitos pelo Gmail. "
        "Confira o histórico e a pasta Enviados antes de repetir; resultados incertos precisam de conferência manual."
        if stage == "envio dos resumos" else
        "A falha ocorreu antes da etapa de envio. Nenhum resumo foi enviado nesta execução."
    )
    message.set_content(
        "Olá! Sou o Hermes, seu assistente pessoal.\n\n"
        "Não consegui concluir a execução dos resumos dos cartões.\n\n"
        f"Cartão: {card.upper() if card else 'não definido'}\n"
        f"Referência: {month or 'não definida'}\n"
        f"Etapa: {stage}\n"
        f"Motivo: {reason}\n\n"
        f"{outcome}\n\n"
        "Este aviso é destinado ao proprietário.\n"
    )
    try:
        connection.send_message(message)
    finally:
        connection.close()
