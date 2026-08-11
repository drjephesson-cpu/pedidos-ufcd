# -*- coding: utf-8 -*-
"""Geração de PDF de pedidos."""

from __future__ import annotations

import io
from datetime import date
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


def _fmt_num(val) -> str:
    if val is None or val == "":
        return "—"
    try:
        n = float(val)
    except (TypeError, ValueError):
        return str(val)
    return str(int(n)) if n == int(n) else str(n)


def _fmt_qtd(val) -> str:
    try:
        n = float(val or 0)
    except (TypeError, ValueError):
        return str(val or "")
    if n == 0:
        return "—"
    return str(int(n)) if n == int(n) else str(n)


def _precisa_pedir(it: dict[str, Any]) -> bool:
    if "pedir" in it and it.get("pedir") is not None:
        return bool(it.get("pedir"))
    try:
        return float(it.get("quantidade") or 0) > 0
    except (TypeError, ValueError):
        return False


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
        pagesize=landscape(A4),
        leftMargin=1.2 * cm,
        rightMargin=1.2 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
        title=titulo,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TituloUFCD",
        parent=styles["Heading1"],
        fontSize=15,
        spaceAfter=4,
        textColor=colors.HexColor("#143d2f"),
    )
    sub_style = ParagraphStyle(
        "SubUFCD",
        parent=styles["Normal"],
        fontSize=9,
        textColor=colors.HexColor("#3d5a4c"),
        spaceAfter=3,
    )
    cell_style = ParagraphStyle(
        "CellUFCD",
        parent=styles["Normal"],
        fontSize=7.5,
        leading=9,
    )
    cell_center = ParagraphStyle(
        "CellCenterUFCD",
        parent=cell_style,
        alignment=1,
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
    story.append(Spacer(1, 0.3 * cm))

    header = [
        "Categoria",
        "Cód.",
        "Medicamento",
        "Est. Mín.",
        "Ponto",
        "Caixa",
        "Estoque",
        "Pedir?",
        "Qtde",
    ]
    data = [header]
    highlight_rows: list[int] = []

    for idx, it in enumerate(itens, start=1):
        pedir = _precisa_pedir(it)
        if pedir:
            highlight_rows.append(idx)
        data.append(
            [
                Paragraph(str(it.get("aba") or "—"), cell_style),
                Paragraph(str(it.get("codigo") or "—"), cell_center),
                Paragraph(str(it.get("descricao") or ""), cell_style),
                Paragraph(_fmt_num(it.get("estoque_minimo")), cell_center),
                Paragraph(_fmt_num(it.get("ponto_pedido")), cell_center),
                Paragraph(_fmt_num(it.get("caixa_com")), cell_center),
                Paragraph(_fmt_num(it.get("estoque_aghu")), cell_center),
                Paragraph("Sim" if pedir else "Não", cell_center),
                Paragraph(_fmt_qtd(it.get("quantidade")), cell_center),
            ]
        )

    if len(data) == 1:
        story.append(Paragraph("Nenhum item no PDF.", styles["Normal"]))
    else:
        # landscape A4 usable ~27.7 cm
        table = Table(
            data,
            colWidths=[
                3.0 * cm,  # categoria
                1.8 * cm,  # cód
                9.2 * cm,  # medicamento
                2.0 * cm,  # est mín
                1.8 * cm,  # ponto
                1.6 * cm,  # caixa
                2.0 * cm,  # estoque
                1.6 * cm,  # pedir
                1.8 * cm,  # qtde
            ],
            repeatRows=1,
        )
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#143d2f")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("FONTSIZE", (0, 1), (-1, -1), 7.5),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#b7c9bf")),
            (
                "ROWBACKGROUNDS",
                (0, 1),
                (-1, -1),
                [colors.white, colors.HexColor("#f3f8f5")],
            ),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]
        highlight = colors.HexColor("#FDEBD0")
        qtde_col = 8  # última coluna: Qtde
        for row in highlight_rows:
            style_cmds.append(("BACKGROUND", (qtde_col, row), (qtde_col, row), highlight))
            style_cmds.append(("FONTNAME", (qtde_col, row), (qtde_col, row), "Helvetica-Bold"))
            style_cmds.append(("TEXTCOLOR", (qtde_col, row), (qtde_col, row), colors.HexColor("#8a3b00")))

        table.setStyle(TableStyle(style_cmds))
        story.append(table)
        total = sum(float(it.get("quantidade") or 0) for it in itens)
        n_pedir = sum(1 for it in itens if _precisa_pedir(it))
        story.append(Spacer(1, 0.3 * cm))
        story.append(
            Paragraph(
                f"<b>{len(itens)}</b> itens · <b>{n_pedir}</b> a pedir · "
                f"quantidade total <b>{int(total) if total == int(total) else total}</b>",
                sub_style,
            )
        )

    doc.build(story)
    return buf.getvalue()
