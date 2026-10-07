"""Cards-specific Gmail authentication policy and controlled diagnostics."""
from pathlib import Path

from automation_core import google_auth
from automation_core.gmail import GmailClient, GmailSendRejected

from .config import valid_email
from .errors import ReportError


def gmail_connect(token_file: Path) -> GmailClient:
    """Use the token selected by application configuration, not library defaults."""
    try:
        connection = GmailClient(google_auth.build_service("gmail", "v1", token_file))
        if not valid_email(connection.email_address):
            connection.close()
            raise ReportError("A conta OAuth do Hermes não retornou um endereço Gmail válido.")
        return connection
    except ReportError:
        raise
    except Exception as exc:
        raise ReportError("Não foi possível autenticar pela conta Google do Hermes. Verifique o token OAuth e o escopo gmail.send.") from exc
