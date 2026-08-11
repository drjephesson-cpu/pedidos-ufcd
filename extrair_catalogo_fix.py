# -*- coding: utf-8 -*-
"""Abre o Excel VISÍVEL. Digite a senha se pedir. Depois extrai o catálogo fixo."""
import json
import shutil
import tempfile
import time
from pathlib import Path

import pythoncom
import win32com.client

SRC = Path(r"C:\Users\jephesson.santos\Downloads\Pedido de MEDICAMENTOS.xlsx")
if not SRC.exists():
    SRC = Path(r"C:\Users\jephesson.santos\Downloads\teste pedidos.xlsx")

OUT = Path(__file__).resolve().parent / "data" / "catalogo.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

tmp = Path(tempfile.gettempdir()) / "pedido_ufcd_extract.xls"
shutil.copy2(SRC, tmp)

print("=" * 60, flush=True)
print("Se o Excel pedir SENHA, digite agora na janela do Excel.", flush=True)
print("Arquivo:", SRC.name, flush=True)
print("=" * 60, flush=True)

pythoncom.CoInitialize()
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(str(tmp), UpdateLinks=0, ReadOnly=True)
print("Aberto!", wb.Sheets.Count, "abas", flush=True)

catalog = {"abas": [], "itens": {}}
for i in range(1, wb.Sheets.Count + 1):
    ws = wb.Sheets(i)
    name = str(ws.Name)
    if name.upper() == "ESTOQUE":
        continue
    vals = ws.UsedRange.Value
    if not vals or not isinstance(vals, tuple):
        continue
    if not isinstance(vals[0], tuple):
        vals = (vals,)
    header_row = None
    for r_idx, row in enumerate(vals[:12]):
        for cell in row[:12] if row else []:
            if isinstance(cell, str):
                vu = cell.upper().replace("Ó", "O").replace("Ê", "E")
                if "AGHU" in vu or "CODIGO" in vu:
                    header_row = r_idx
                    break
        if header_row is not None:
            break
    if header_row is None:
        continue
    items = []
    for row in vals[header_row + 1 :]:
        code = row[0]
        if code is None or isinstance(code, str):
            continue
        try:
            code_int = int(code)
        except Exception:
            continue
        desc = row[1] if len(row) > 1 else ""
        est_min = row[2] if len(row) > 2 else None
        ponto = row[3] if len(row) > 3 else None
        caixa = row[4] if len(row) > 4 else None
        if est_min is None and ponto is None:
            continue
        try:
            items.append(
                {
                    "codigo": code_int,
                    "descricao": str(desc).strip() if desc else "",
                    "estoque_minimo": float(est_min or 0),
                    "ponto_pedido": float(ponto or 0),
                    "caixa_com": float(caixa) if caixa not in (None, 0, "") else 1.0,
                }
            )
        except Exception:
            continue
    catalog["abas"].append({"id": name, "titulo": name, "count": len(items)})
    catalog["itens"][name] = items
    print(f"  {name}: {len(items)} itens", flush=True)

wb.Close(False)
excel.Quit()
pythoncom.CoUninitialize()

OUT.write_text(json.dumps(catalog, ensure_ascii=False, indent=2), encoding="utf-8")
total = sum(a["count"] for a in catalog["abas"])
print(f"\nCatálogo fixo salvo: {OUT}", flush=True)
print(f"{len(catalog['abas'])} abas, {total} itens", flush=True)
try:
    tmp.unlink(missing_ok=True)
except Exception:
    pass
