"""Controlled operational diagnostics safe for CLI output and owner alerts."""

from __future__ import annotations


class ReportError(Exception):
    """Erro operacional legível, sem credenciais ou conteúdo dos gastos."""
