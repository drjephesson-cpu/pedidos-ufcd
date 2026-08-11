# -*- coding: utf-8 -*-
"""Geração de PDF de pedidos."""

from __future__ import annotations

import io
from datetime import date
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


def _fmt_qtd(val) -> str:
    try:
        n = float(val or 0)
    except (TypeError, ValueError):
        return str(val or "")
    return str(int(n)) if n == int(n) else str(n)


def gerar_pdf_pedido(
    *,
    titulo: str,
    data_pedido: date | str,
    usuario: str | None,
    itens: list[dict[str, Any]],
    observacao: str | None = None,
) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
        title=titulo,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TituloUFCD",
        parent=styles["Heading1"],
        fontSize=16,
        spaceAfter=6,
        textColor=colors.HexColor("#143d2f"),
    )
    sub_style = ParagraphStyle(
        "SubUFCD",
        parent=styles["Normal"],
        fontSize=10,
        textColor=colors.HexColor("#3d5a4c"),
        spaceAfter=4,
    )
    cell_style = ParagraphStyle(
        "CellUFCD",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
    )

    story = []
    story.append(Paragraph("Pedidos UFCD — Farmácia", title_style))
    story.append(Paragraph(titulo, sub_style))
    story.append(
        Paragraph(
            f"Data do pedido: <b>{data_pedido}</b>"
            + (f" · Usuário: {usuario}" if usuario else ""),
            sub_style,
        )
    )
    if observacao:
        story.append(Paragraph(f"Obs.: {observacao}", sub_style))
    story.append(Spacer(1, 0.4 * cm))

    header = ["Aba", "Cód.", "Medicamento", "Qtde", "Origem"]
    data = [header]
    for it in itens:
        data.append(
            [
                Paragraph(str(it.get("aba") or "—"), cell_style),
                str(it.get("codigo") or "—"),
                Paragraph(str(it.get("descricao") or ""), cell_style),
                f"{_fmt_qtd(it.get('quantidade'))}",
                "Manual" if it.get("origem") == "manual" else "Auto",
            ]
        )

    if len(data) == 1:
        story.append(Paragraph("Nenhum item a pedir.", styles["Normal"]))
    else:
        table = Table(data, colWidths=[3.2 * cm, 2.0 * cm, 9.0 * cm, 1.6 * cm, 1.8 * cm])
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#143d2f")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 9),
                    ("FONTSIZE", (0, 1), (-1, -1), 8),
                    ("ALIGN", (1, 0), (1, -1), "CENTER"),
                    ("ALIGN", (3, 0), (4, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#b7c9bf")),
                    (
                        "ROWBACKGROUNDS",
                        (0, 1),
                        (-1, -1),
                        [colors.white, colors.HexColor("#f3f8f5")],
                    ),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(table)
        total = sum(float(it.get("quantidade") or 0) for it in itens)
        story.append(Spacer(1, 0.35 * cm))
        story.append(
            Paragraph(
                f"<b>{len(itens)}</b> itens · quantidade total <b>{int(total) if total == int(total) else total}</b>",
                sub_style,
            )
        )

    doc.build(story)
    return buf.getvalue()
