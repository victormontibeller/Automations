"""Pure HTML and plain-text rendering of individual card reports."""

from __future__ import annotations

from decimal import Decimal
from html import escape

from .models import ZERO, Report

BLUE = "#2F75B5"


def brl(value: Decimal) -> str:
    formatted = f"{abs(value):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return ("-" if value < ZERO else "") + "R$ " + formatted


def subject(report: Report) -> str:
    return f"Cartão {report.card.upper()} {report.month} — {report.recipient.sheet}"


def greeting_paragraphs(report: Report) -> tuple[str, str, str, str]:
    owner = report.owner_name or "[seu nome]"
    year, month = report.month.split("-")
    return (
        f"Olá, {report.recipient.sheet}! Tudo bem?",
        f"Sou o Hermes, assistente pessoal do {owner}. Vou ajudar no envio dos resumos mensais dos cartões compartilhados.",
        f"Segue abaixo o seu resumo, referente a {month}/{year}, com os gastos, pagamentos registrados e saldo atualizado.",
        f"Se tiver alguma dúvida sobre os lançamentos, fale diretamente com {owner}.",
    )


def render_html(report: Report) -> str:
    """Use inline table styles supported by Gmail, without scripts or remote assets."""
    number = "text-align:right;white-space:nowrap;"
    cell_style = "padding:7px 6px;"
    row_open = '<tr style="border-bottom:1px solid #e5e7eb;vertical-align:top;">'

    def cell(text: str, style: str = "", colspan: int = 1) -> str:
        span = f' colspan="{colspan}"' if colspan > 1 else ""
        return f'<td{span} style="{cell_style}{style}">{escape(text)}</td>'

    rows = []
    for p in report.purchases:
        rows.append(row_open + cell(p.date.strftime("%d/%m/%Y"), "white-space:nowrap;")
                    + cell(p.description, "overflow-wrap:anywhere;word-break:break-word;")
                    + cell(p.installment) + cell(brl(p.total), number)
                    + cell(brl(p.share), number) + "</tr>")

    def summary_row(label: str, amount: Decimal, strong: bool = False) -> str:
        style = "font-weight:bold;" if strong else ""
        if strong:
            style += "border-top:2px solid #9ca3af;"
        # Leave the original purchase Total column empty in summary rows.
        return row_open + cell(label, style, 3) + cell("", style) + cell(brl(amount), number + style) + "</tr>"

    rows.append(summary_row("TOTAL", report.spending, True))
    rows.extend(summary_row("PAGAMENTOS", amount) for amount in (report.payments or (ZERO,)))
    rows.append(summary_row("TOTAL GERAL", report.balance, True))
    status = "Saldo quitado." if report.balance == ZERO else "Crédito a seu favor." if report.balance < ZERO else ""
    status_paragraph = f'<p style="font-size:14px;margin:12px 0 4px;">{status}</p>' if status else ""
    title = f"Cartão {report.card.upper()} {report.month}"
    heads = "".join(
        f'<th scope="col" style="padding:10px 6px;font-size:14px;text-align:{alignment};border-bottom:1px solid #d1d5db;">{label}</th>'
        for label, alignment in (("Data", "left"), ("Lançamento", "left"), ("Parcelas", "left"), ("Total", "right"), ("Rateio", "right"))
    )
    greeting = greeting_paragraphs(report)
    introduction = "".join(
        f'<p style="font-size:16px;line-height:1.5;margin:8px 0 14px;">{escape(paragraph)}</p>'
        for paragraph in greeting[:3]
    )
    payment_footer = (
        f'<p style="font-size:16px;line-height:1.5;margin:14px 0 8px;">{escape(report.payment_footer)}</p>'
        if report.payment_footer else ""
    )
    return f'''<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{escape(subject(report))}</title></head>
<body style="margin:0;padding:12px;background:#ffffff;color:#111111;font-family:Calibri,Arial,sans-serif;">
<table role="presentation" width="100%" style="width:100%;max-width:920px;margin:0 auto;border-collapse:collapse;"><tr><td>
{introduction}
<table width="100%" aria-label="Gastos do cartão" style="width:100%;border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:14px;">
<thead><tr><th colspan="5" bgcolor="{BLUE}" style="background:{BLUE};color:#111111;text-align:center;padding:12px;font-size:24px;font-weight:bold;">{escape(title)}</th></tr><tr>{heads}</tr></thead>
<tbody>{''.join(rows)}</tbody></table>
{status_paragraph}
<p style="font-size:16px;line-height:1.5;margin:14px 0 8px;">{escape(greeting[3])}</p>
{payment_footer}
</td></tr></table></body></html>'''


def render_text(report: Report) -> str:
    greeting = greeting_paragraphs(report)
    lines = ["\n\n".join(greeting[:3]), "", subject(report), "", "Data | Lançamento | Parcelas | Total | Rateio"]
    lines += [f"{p.date:%d/%m/%Y} | {p.description} | {p.installment} | {brl(p.total)} | {brl(p.share)}" for p in report.purchases]
    lines += ["", "TOTAL: " + brl(report.spending)]
    lines += ["PAGAMENTOS: " + brl(p) for p in (report.payments or (ZERO,))]
    lines += ["TOTAL GERAL: " + brl(report.balance)]
    lines += ["", greeting[3]]
    if report.payment_footer:
        lines += ["", report.payment_footer]
    return "\n".join(lines) + "\n"
