"""Application orchestration and controlled failure/owner-alert policy."""

from __future__ import annotations

import hashlib
from io import BytesIO
from pathlib import Path
import sys

from . import alerts, config, delivery, drive, ledger, previews, scheduling, workbook
from .config import Settings
from .errors import ReportError


def input_workbook_bytes(settings: Settings, month: str) -> bytes:
    if settings.input_source == "drive":
        return drive.drive_workbook_bytes(settings, month)
    filename = f"{month}.xlsx"
    try:
        return (settings.input_dir / filename).read_bytes()
    except OSError as exc:
        raise ReportError(f"Arquivo {filename} ausente ou inacessível na pasta local configurada.") from exc


def run(config_path: Path, *, card: str | None = None,
        requested_month: str | None = None, recipient: str | None = None,
        send: bool = False, history: bool = False, resolve: str | None = None,
        test_to: str | None = None, scheduled: bool = False) -> int:
    """Run one request; previews/history never connect to Gmail.

    Stage names and exit codes are part of the existing operational contract.
    Input bytes are snapshotted once for both parsing and the delivery hash.
    """
    settings = None
    month = None
    stage = "configuração"
    try:
        settings = config.load_settings(config_path)
        if requested_month:
            month = scheduling.validate_month(requested_month)
        if history or resolve:
            if scheduled or test_to:
                raise ReportError("Histórico não aceita --scheduled ou --test-to.")
            ledger.maintain_ledger(settings, requested_month, card, recipient, resolve)
            return 0
        if not card:
            raise ReportError("Informe --card black ou --card latam.")
        if scheduled:
            if requested_month:
                raise ReportError("--scheduled não aceita --month; recuperação deve ser manual.")
            month = scheduling.scheduled_month(card)
        else:
            if not requested_month:
                raise ReportError("Informe --month AAAA-MM para executar manualmente.")
            month = requested_month
        if test_to and not send:
            raise ReportError("--test-to exige --send.")
        stage = "busca e validação da planilha"
        # Snapshot the bytes once so the audit hash and report share the same input.
        source_bytes = input_workbook_bytes(settings, month)
        # read_reports accepts a stream as well as a path; preserve the expected filename.
        stream = BytesIO(source_bytes)
        stream.name = f"{month}.xlsx"
        reports = workbook.read_reports(stream, settings, card, month)
        if recipient:
            reports = [r for r in reports if r.recipient.sheet == recipient]
            if not reports:
                raise ReportError("Aba pessoal não cadastrada ou desativada.")
        active = [r for r in reports if r.has_movement]
        if send:
            stage = "envio dos resumos"
            count, skipped = delivery.send_reports(active, settings, hashlib.sha256(source_bytes).hexdigest(), test_to=test_to)
            print(f"{card.upper()} {month}: {count} mensagens {'de teste ' if test_to else ''}aceitas pela API do Gmail; {skipped} já enviadas; {len(reports) - len(active)} sem movimento.")
        else:
            folder = settings.output_dir / month / card
            previews.write_previews(active, folder)
            print(f"Prévia {card.upper()} {month}: {len(active)} resumos em {folder}; {len(reports) - len(active)} sem movimento. Nenhum e-mail enviado.")
        return 0
    except Exception as exc:
        # Only our controlled diagnostics may include details. Raw exceptions
        # can contain cell values, Gmail addresses or credential server output.
        reason = str(exc) if isinstance(exc, ReportError) else f"Falha operacional ({type(exc).__name__}). Verifique Gmail, permissões e acesso aos arquivos e ao histórico."
        print(f"Erro: {reason}", file=sys.stderr)
        if send:
            try:
                alerts.notify_owner(settings, card, month, stage, reason)
            except Exception as alert_error:
                # Never retry recursively, notify friends or mask the failure.
                print(f"Não foi possível confirmar o aviso ao proprietário ({type(alert_error).__name__}). Confira o Gmail e os logs do Hermes.", file=sys.stderr)
            else:
                print("Aviso de falha aceito pelo Gmail para o proprietário (Cc pessoal, se configurado).", file=sys.stderr)
        return 1
