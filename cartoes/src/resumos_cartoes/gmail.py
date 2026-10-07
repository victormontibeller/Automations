"""Local Gmail OAuth adapter; sends are never automatically retried."""

from __future__ import annotations

import base64
from email.message import EmailMessage
from pathlib import Path

from .config import valid_email
from .errors import ReportError


class GmailSendRejected(Exception):
    """The Gmail API explicitly rejected a message before accepting it."""


class GmailAPIConnection:
    def __init__(self, service):
        self.service = service
        try:
            profile = service.users().getProfile(userId="me").execute(num_retries=0)
            self.email_address = profile.get("emailAddress", "")
            if not valid_email(self.email_address):
                raise ReportError("A conta OAuth do Hermes não retornou um endereço Gmail válido.")
        except Exception:
            self.close()
            raise

    def send_message(self, message: EmailMessage):
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("ascii")
        try:
            return self.service.users().messages().send(
                userId="me", body={"raw": raw}
            ).execute(num_retries=0)
        except Exception as exc:
            status = getattr(getattr(exc, "resp", None), "status", None)
            if isinstance(status, int) and 400 <= status < 500:
                raise GmailSendRejected("Gmail API recusou a mensagem.") from exc
            raise

    def close(self):
        http = getattr(self.service, "_http", None)
        close = getattr(http, "close", None)
        if close:
            close()


def _build_gmail_service(token_file: Path):
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    credentials = Credentials.from_authorized_user_file(str(token_file))
    return build("gmail", "v1", credentials=credentials, cache_discovery=False)


def gmail_connect(token_file: Path):
    """Build a Gmail API client from the existing Hermes OAuth token."""
    try:
        return GmailAPIConnection(_build_gmail_service(token_file))
    except ReportError:
        raise
    except Exception as exc:
        raise ReportError("Não foi possível autenticar pela conta Google do Hermes. Verifique o token OAuth e o escopo gmail.send.") from exc
