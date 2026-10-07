"""CLI argument parsing; execution policy belongs to the application service."""

from __future__ import annotations

import argparse
from pathlib import Path

from . import service
from .scheduling import CARDS


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Resumos individuais dos cartões. Sem --send, apenas gera prévias.")
    parser.add_argument("--config", type=Path, default=Path("config.json"))
    parser.add_argument("--card", choices=CARDS)
    parser.add_argument("--month", help="AAAA-MM; obrigatório em execução manual")
    parser.add_argument("--recipient", help="Restringir a uma aba pessoal, pelo nome exato")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--dry-run", action="store_true", help="Gerar prévias, sem Gmail (padrão)")
    mode.add_argument("--send", action="store_true", help="Enviar mensagens")
    mode.add_argument("--history", action="store_true", help="Consultar o histórico")
    mode.add_argument("--resolve", choices=("sent", "not-sent"), help="Resolver resultado incerto após conferir o Gmail")
    parser.add_argument("--test-to", help="Com --send, redirecionar todos os resumos a um endereço de teste")
    parser.add_argument("--scheduled", action="store_true", help="Validar dia previsto. Black dia 5 = mês anterior; Latam dia 20 = mês atual")
    args = parser.parse_args(argv)
    return service.run(
        args.config, card=args.card, requested_month=args.month,
        recipient=args.recipient, send=args.send, history=args.history,
        resolve=args.resolve, test_to=args.test_to, scheduled=args.scheduled,
    )
