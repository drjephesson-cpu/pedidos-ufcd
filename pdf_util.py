# -*- coding: utf-8 -*-
"""Geração de PDF de pedidos."""

from __future__ import annotations

import io
from collections import OrderedDict
from datetime import date
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    PageBreak,
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


def _agrupar_por_categoria(itens: list[dict[str, Any]]) -> OrderedDict[str, list[dict]]:
    grupos: OrderedDict[str, list[dict]] = OrderedDict()
    for it in itens:
        aba = str(it.get("aba") or "Sem categoria").strip() or "Sem categoria"
        grupos.setdefault(aba, []).append(it)
    return grupos


def _montar_tabela(
    itens: list[dict[str, Any]],
    *,
    cell_style: ParagraphStyle,
    cell_center: ParagraphStyle,
    mostrar_categoria: bool,
) -> Table | None:
    if not itens:
        return None

    if mostrar_categoria:
        header = [
            "Categoria",
            "Cód.",
            "Medicamento",
            "Est. Mín.",
            "Ponto",
            "Caixa",
            "Estoque",
            "Qtde",
        ]
        col_widths = [
            3.0 * cm,
            1.8 * cm,
            11.0 * cm,
            2.0 * cm,
            1.8 * cm,
            1.6 * cm,
            2.0 * cm,
            2.0 * cm,
        ]
        qtde_col = 7
    else:
        header = [
            "Cód.",
            "Medicamento",
            "Est. Mín.",
            "Ponto",
            "Caixa",
            "Estoque",
            "Qtde",
        ]
        col_widths = [
            2.0 * cm,
            14.0 * cm,
            2.0 * cm,
            1.8 * cm,
            1.6 * cm,
            2.0 * cm,
            2.0 * cm,
        ]
        qtde_col = 6

    data = [header]
    highlight_rows: list[int] = []

    for idx, it in enumerate(itens, start=1):
        if _precisa_pedir(it):
            highlight_rows.append(idx)
        row = []
        if mostrar_categoria:
            row.append(Paragraph(str(it.get("aba") or "—"), cell_style))
        row.extend(
            [
                Paragraph(str(it.get("codigo") or "—"), cell_center),
                Paragraph(str(it.get("descricao") or ""), cell_style),
                Paragraph(_fmt_num(it.get("estoque_minimo")), cell_center),
                Paragraph(_fmt_num(it.get("ponto_pedido")), cell_center),
                Paragraph(_fmt_num(it.get("caixa_com")), cell_center),
                Paragraph(_fmt_num(it.get("estoque_aghu")), cell_center),
                Paragraph(_fmt_qtd(it.get("quantidade")), cell_center),
            ]
        )
        data.append(row)

    table = Table(data, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.black),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTSIZE", (0, 1), (-1, -1), 7.5),
        ("TEXTCOLOR", (0, 1), (-1, -1), colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
        (
            "ROWBACKGROUNDS",
            (0, 1),
            (-1, -1),
            [colors.white, colors.Color(0.92, 0.92, 0.92)],
        ),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]
    for row in highlight_rows:
        style_cmds.append(
            ("BACKGROUND", (qtde_col, row), (qtde_col, row), colors.Color(0.85, 0.85, 0.85))
        )
        style_cmds.append(("FONTNAME", (qtde_col, row), (qtde_col, row), "Helvetica-Bold"))
        style_cmds.append(("TEXTCOLOR", (qtde_col, row), (qtde_col, row), colors.black))

    table.setStyle(TableStyle(style_cmds))
    return table


def _resumo_paragrafo(itens: list[dict[str, Any]], sub_style: ParagraphStyle) -> Paragraph:
    total = sum(float(it.get("quantidade") or 0) for it in itens)
    n_pedir = sum(1 for it in itens if _precisa_pedir(it))
    total_txt = int(total) if total == int(total) else total
    return Paragraph(
        f"<b>{len(itens)}</b> itens · <b>{n_pedir}</b> a pedir · "
        f"quantidade total <b>{total_txt}</b>",
        sub_style,
    )


def gerar_pdf_pedido(
    *,
    titulo: str,
    data_pedido: date | str,
    usuario: str | None,
    itens: list[dict[str, Any]],
    observacao: str | None = None,
    separar_paginas: bool = False,
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
        textColor=colors.black,
    )
    sub_style = ParagraphStyle(
        "SubUFCD",
        parent=styles["Normal"],
        fontSize=9,
        textColor=colors.black,
        spaceAfter=3,
    )
    cat_style = ParagraphStyle(
        "CatUFCD",
        parent=styles["Heading2"],
        fontSize=12,
        spaceBefore=2,
        spaceAfter=6,
        textColor=colors.black,
    )
    cell_style = ParagraphStyle(
        "CellUFCD",
        parent=styles["Normal"],
        fontSize=7.5,
        leading=9,
        textColor=colors.black,
    )
    cell_center = ParagraphStyle(
        "CellCenterUFCD",
        parent=cell_style,
        alignment=1,
    )

    def cabecalho(extra: str | None = None) -> list:
        blocos = [
            Paragraph("Pedidos UFCD — Farmácia", title_style),
            Paragraph(titulo, sub_style),
            Paragraph(
                f"Data do pedido: <b>{data_pedido}</b>"
                + (f" · Usuário: {usuario}" if usuario else ""),
                sub_style,
            ),
        ]
        if observacao:
            blocos.append(Paragraph(f"Obs.: {observacao}", sub_style))
        if extra:
            blocos.append(Paragraph(extra, cat_style))
        blocos.append(Spacer(1, 0.25 * cm))
        return blocos

    story: list = []

    if not itens:
        story.extend(cabecalho())
        story.append(Paragraph("Nenhum item no PDF.", styles["Normal"]))
        doc.build(story)
        return buf.getvalue()

    grupos = _agrupar_por_categoria(itens)
    usar_paginas = separar_paginas and len(grupos) > 1

    if not usar_paginas:
        story.extend(cabecalho())
        tabela = _montar_tabela(
            itens,
            cell_style=cell_style,
            cell_center=cell_center,
            mostrar_categoria=True,
        )
        if tabela is not None:
            story.append(tabela)
            story.append(Spacer(1, 0.3 * cm))
            story.append(_resumo_paragrafo(itens, sub_style))
    else:
        for i, (aba, grupo) in enumerate(grupos.items()):
            if i > 0:
                story.append(PageBreak())
            story.extend(cabecalho(extra=f"Categoria: {aba}"))
            tabela = _montar_tabela(
                grupo,
                cell_style=cell_style,
                cell_center=cell_center,
                mostrar_categoria=False,
            )
            if tabela is not None:
                story.append(tabela)
                story.append(Spacer(1, 0.3 * cm))
                story.append(_resumo_paragrafo(grupo, sub_style))

    doc.build(story)
    return buf.getvalue()
