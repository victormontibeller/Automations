"""Application configuration, mailbox validation and Hermes token path policy."""

from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
import re

from .errors import ReportError
from .models import Recipient

DEFAULT_SENDER_NAME = "Hermes"


@dataclass(frozen=True)
class Settings:
    input_dir: Path
    input_source: str
    drive_folder_name: str
    output_dir: Path
    state_dir: Path
    google_token_file: Path
    sender_name: str
    recipients: tuple[Recipient, ...]
    owner_name: str | None = None
    payment_footer: str | None = None
    personal_copy_email: str | None = None


def _clean_string(value: object, label: str) -> str:
    if not isinstance(value, str) or not value.strip() or any(c in value for c in "\r\n\x00"):
        raise ReportError(f"{label}: informe um texto válido em uma linha.")
    return value.strip()


def valid_email(value: str | None) -> bool:
    # One plain mailbox only: no display names, lists, BCC or header injection.
    if not isinstance(value, str) or len(value) > 254 or value.count("@") != 1:
        return False
    local, domain = value.split("@")
    if not re.fullmatch(r"[A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]{1,64}", local):
        return False
    if local.startswith(".") or local.endswith(".") or ".." in local:
        return False
    labels = domain.split(".")
    return len(labels) >= 2 and all(re.fullmatch(r"[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?", label) for label in labels)


def hermes_home() -> Path:
    configured = os.environ.get("HERMES_HOME")
    return Path(configured).expanduser() if configured else Path.home() / ".hermes"


def load_settings(path: Path) -> Settings:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ReportError("Não foi possível ler o arquivo de configuração JSON.") from exc
    if not isinstance(data, dict) or not isinstance(data.get("recipients"), list):
        raise ReportError("Configuração deve conter uma lista recipients.")
    recipients = []
    for item in data["recipients"]:
        if not isinstance(item, dict):
            raise ReportError("Cada destinatário deve ser um objeto JSON.")
        sheet = _clean_string(item.get("sheet"), "sheet")
        participant = _clean_string(item.get("participant"), "participant")
        email = item.get("email")
        if email is not None and not valid_email(email):
            raise ReportError(f"E-mail inválido no cadastro de {sheet}.")
        enabled = item.get("enabled", True)
        if not isinstance(enabled, bool):
            raise ReportError(f"enabled deve ser true ou false em {sheet}.")
        recipients.append(Recipient(sheet, participant, email, enabled))
    if not recipients:
        raise ReportError("Cadastre pelo menos uma aba pessoal.")
    for attr in ("sheet", "participant"):
        keys = [getattr(r, attr).casefold() for r in recipients]
        if len(keys) != len(set(keys)):
            raise ReportError(f"Cadastro duplicado no campo {attr}.")
    if any(r.sheet in ("Black", "Latam") for r in recipients):
        raise ReportError("As abas dos cartões não são destinatários.")
    input_source = data.get("input_source", "drive")
    if not isinstance(input_source, str) or input_source not in {"drive", "local"}:
        raise ReportError("input_source deve ser 'drive' ou 'local'.")
    drive_folder_name = _clean_string(data.get("drive_folder_name", "Cartão"), "drive_folder_name")
    base = path.resolve().parent

    def local_path(key: str, default: str) -> Path:
        value = _clean_string(data.get(key, default), key)
        return (base / Path(value).expanduser()).resolve()

    owner_name = data.get("owner_name")
    if owner_name is not None:
        owner_name = _clean_string(owner_name, "owner_name")
    payment_footer = data.get("payment_footer")
    if payment_footer is not None:
        payment_footer = _clean_string(payment_footer, "payment_footer")
    personal_copy_email = data.get("personal_copy_email")
    if personal_copy_email is not None and not valid_email(personal_copy_email):
        raise ReportError("personal_copy_email deve conter um único e-mail válido.")
    return Settings(
        local_path("input_dir", "inputs"), input_source, drive_folder_name,
        local_path("output_dir", "outputs"), local_path("state_dir", "var"),
        local_path("google_token_file", str(hermes_home() / "google_token.json")),
        _clean_string(data.get("sender_name", DEFAULT_SENDER_NAME), "sender_name"),
        tuple(recipients), owner_name, payment_footer, personal_copy_email,
    )
