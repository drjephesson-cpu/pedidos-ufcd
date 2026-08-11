# -*- coding: utf-8 -*-
"""Pedido de Estoque — Farmácia UFCD

Replica a lógica da planilha 'teste pedidos':
- Match por Cód. AGHU
- Pedir? = ponto_pedido > estoque
- Quanto Pedir? = CEILING(estoque_minimo - estoque, caixa_com)
  (ou valor manual, se informado)
"""

from __future__ import annotations

import io
import json
import math
import re
import tempfile
from datetime import date
from pathlib import Path

from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from werkzeug.utils import secure_filename

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)
CATALOGO_PATH = DATA / "catalogo.json"
ESTOQUE_PATH = DATA / "estoque_atual.json"

app = Flask(__name__)
app.secret_key = "pedido-estoque-ufcd-local"

# Abas na mesma ordem da planilha (exceto Estoque)
ABA_ORDEM_PADRAO = [
    "COMPRIMIDOS",
    "SOLUÇÕES",
    "Geladeira",
    "PÓS",
    "CREMES e Pomadas",
    "M.P.P.S",
    "SOROS",
    "LÍQUIDOS E GOTAS",
    "AMPOLAS",
    "RESERVA TERPÊUTICA",
    "ALMOXARIFADO",
]


def ceiling_math(value: float, significance: float) -> float:
    """Equivalente ao CEILING.MATH do Excel."""
    if significance in (0, None):
        significance = 1
    if value <= 0:
        return 0
    return math.ceil(value / significance) * significance


def load_catalogo() -> dict:
    if not CATALOGO_PATH.exists():
        return {"abas": [], "itens": {}}
    return json.loads(CATALOGO_PATH.read_text(encoding="utf-8"))


def save_catalogo(catalogo: dict) -> None:
    CATALOGO_PATH.write_text(
        json.dumps(catalogo, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def load_estoque() -> dict:
    """Mapa codigo(str) -> saldo (float)."""
    if not ESTOQUE_PATH.exists():
        return {}
    return json.loads(ESTOQUE_PATH.read_text(encoding="utf-8"))


def save_estoque(mapa: dict, meta: dict | None = None) -> None:
    payload = {"saldos": mapa, "meta": meta or {}}
    ESTOQUE_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def estoque_saldos() -> dict:
    raw = load_estoque()
    if isinstance(raw, dict) and "saldos" in raw:
        return {str(k): v for k, v in raw["saldos"].items()}
    return {str(k): v for k, v in raw.items()}


def estoque_meta() -> dict:
    raw = load_estoque()
    if isinstance(raw, dict) and "meta" in raw:
        return raw["meta"]
    return {}


def normaliza_codigo(val) -> str | None:
    if val is None or val == "":
        return None
    if isinstance(val, str):
        val = val.strip()
        if not val or not re.match(r"^-?\d+(\.0+)?$", val):
            # tenta número puro
            try:
                return str(int(float(val.replace(",", "."))))
            except Exception:
                return None
        try:
            return str(int(float(val)))
        except Exception:
            return None
    try:
        return str(int(float(val)))
    except Exception:
        return None


def parse_estoque_xlsx(file_storage) -> tuple[dict, dict]:
    """Lê EstoqueFarmacia: colunas codigo + saldo_farmacia_central (ou 8ª col)."""
    wb = load_workbook(file_storage, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        return {}, {}

    header_l = [str(h).strip().lower() if h is not None else "" for h in header]
    cod_idx = None
    saldo_idx = None
    for i, h in enumerate(header_l):
        if h in ("codigo", "código", "cod_aghu", "cód. aghu", "cod. aghu"):
            cod_idx = i
        if h in (
            "saldo_farmacia_central",
            "saldo",
            "estoque",
            "estoque_aghu",
            "qtde",
            "quantidade",
        ):
            saldo_idx = i

    # fallback: 1ª coluna = código, 8ª = saldo (como na planilha)
    if cod_idx is None:
        cod_idx = 0
    if saldo_idx is None:
        saldo_idx = 7 if len(header) >= 8 else len(header) - 1

    mapa: dict[str, float | None] = {}
    for row in rows:
        if not row:
            continue
        cod = normaliza_codigo(row[cod_idx] if cod_idx < len(row) else None)
        if not cod:
            continue
        saldo_raw = row[saldo_idx] if saldo_idx < len(row) else None
        if saldo_raw is None or saldo_raw == "":
            mapa[cod] = None  # sem saldo informado
        else:
            try:
                mapa[cod] = float(saldo_raw)
            except Exception:
                mapa[cod] = None

    arquivo = getattr(file_storage, "filename", None) or getattr(
        file_storage, "name", None
    )
    if arquivo:
        arquivo = Path(str(arquivo)).name
    meta = {
        "arquivo": arquivo or "",
        "itens": len(mapa),
        "importado_em": date.today().isoformat(),
    }
    return mapa, meta


def _is_header_cell(val) -> bool:
    if not isinstance(val, str):
        return False
    vu = val.upper().replace("Ó", "O").replace("Ê", "E")
    return "AGHU" in vu or "CODIGO" in vu


def parse_pedido_workbook(wb) -> dict:
    """Extrai abas de pedido de um workbook openpyxl."""
    catalog: dict = {"abas": [], "itens": {}}
    for name in wb.sheetnames:
        if name.upper() == "ESTOQUE":
            continue
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        header_row = None
        for r_idx, row in enumerate(rows[:12]):
            for cell in row[:12]:
                if _is_header_cell(cell):
                    header_row = r_idx
                    break
            if header_row is not None:
                break
        if header_row is None:
            continue
        items = []
        for row in rows[header_row + 1 :]:
            if not row:
                continue
            cod = normaliza_codigo(row[0])
            if not cod:
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
                        "codigo": int(cod),
                        "descricao": str(desc).strip() if desc else "",
                        "estoque_minimo": float(est_min or 0),
                        "ponto_pedido": float(ponto or 0),
                        "caixa_com": float(caixa)
                        if caixa not in (None, 0, "")
                        else 1.0,
                    }
                )
            except Exception:
                continue
        catalog["abas"].append({"id": name, "titulo": name, "count": len(items)})
        catalog["itens"][name] = items
    return catalog


def import_pedido_file(file_bytes: bytes, filename: str, password: str | None) -> dict:
    """Importa catálogo de .xlsx/.xls, com suporte a senha (msoffcrypto)."""
    name_lower = (filename or "").lower()
    data = file_bytes

    # Tenta descriptografar se necessário
    try:
        import msoffcrypto

        bio = io.BytesIO(file_bytes)
        office = msoffcrypto.OfficeFile(bio)
        if office.is_encrypted():
            if not password:
                raise ValueError(
                    "Este arquivo está protegido por senha. Informe a senha para importar."
                )
            bio.seek(0)
            office = msoffcrypto.OfficeFile(bio)
            office.load_key(password=password)
            dec = io.BytesIO()
            office.decrypt(dec)
            data = dec.getvalue()
    except ValueError:
        raise
    except Exception as e:
        # se não for encrypted / lib falhar, segue com bytes originais
        if "password" in str(e).lower() or "decrypt" in str(e).lower():
            raise ValueError(f"Não foi possível abrir com a senha informada: {e}") from e

    # Detecta formato
    if data[:2] == b"PK":
        wb = load_workbook(io.BytesIO(data), data_only=True)
        return parse_pedido_workbook(wb)

    # OLE legado: tenta via Excel COM
    if data[:4] == b"\xd0\xcf\x11\xe0":
        return _import_via_excel_com(data, password)

    raise ValueError(
        "Formato não reconhecido. Salve a planilha como .xlsx (Excel) e tente de novo."
    )


def _import_via_excel_com(data: bytes, password: str | None) -> dict:
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise ValueError(
            "Arquivo legado (.xls criptografado). Salve como .xlsx no Excel ou "
            "instale pywin32 para importação automática."
        ) from e

    pythoncom.CoInitialize()
    tmp = Path(tempfile.gettempdir()) / "pedido_import_tmp.xls"
    tmp.write_bytes(data)
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        kwargs = {"UpdateLinks": 0, "ReadOnly": True}
        if password:
            wb = excel.Workbooks.Open(str(tmp), Password=password, **kwargs)
        else:
            wb = excel.Workbooks.Open(str(tmp), **kwargs)

        # Salva temporário xlsx e lê com openpyxl
        xlsx_tmp = Path(tempfile.gettempdir()) / "pedido_import_tmp.xlsx"
        wb.SaveAs(str(xlsx_tmp), FileFormat=51)
        wb.Close(False)
        wb = None
        catalog = parse_pedido_workbook(load_workbook(xlsx_tmp, data_only=True))
        try:
            xlsx_tmp.unlink(missing_ok=True)
        except Exception:
            pass
        return catalog
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass


def calcular_item(item: dict, saldo, manual=None) -> dict:
    est_min = float(item["estoque_minimo"])
    ponto = float(item["ponto_pedido"])
    caixa = float(item["caixa_com"] or 1)
    if caixa <= 0:
        caixa = 1

    estoque = None if saldo is None else float(saldo)
    pedir = False
    quanto = None

    if estoque is not None:
        pedir = ponto > estoque
        if manual not in (None, ""):
            try:
                quanto = float(manual)
            except Exception:
                quanto = None
        elif pedir:
            quanto = ceiling_math(est_min - estoque, caixa)

    return {
        **item,
        "estoque_aghu": estoque,
        "pedir": pedir,
        "quanto_pedir": quanto,
        "manual": manual if manual not in (None, "") else None,
        "sem_estoque": estoque is None,
    }


def montar_aba(aba_id: str, catalogo: dict, saldos: dict, manuais: dict) -> list:
    itens = catalogo.get("itens", {}).get(aba_id, [])
    result = []
    for item in itens:
        cod = str(item["codigo"])
        saldo = saldos.get(cod)
        # se chave não existe, sem_estoque; se existe com None, também
        if cod not in saldos:
            saldo = None
        manual = manuais.get(cod)
        result.append(calcular_item(item, saldo, manual))
    return result


@app.route("/")
def index():
    catalogo = load_catalogo()
    saldos = estoque_saldos()
    meta = estoque_meta()
    abas = catalogo.get("abas") or []
    # ordenar como na planilha quando possível
    ordem = {n: i for i, n in enumerate(ABA_ORDEM_PADRAO)}
    abas_sorted = sorted(
        abas, key=lambda a: ordem.get(a["id"], 100 + abas.index(a))
    )

    aba_atual = request.args.get("aba") or (abas_sorted[0]["id"] if abas_sorted else None)
    filtro = request.args.get("filtro", "todos")  # todos | pedir | sem
    q = (request.args.get("q") or "").strip().lower()

    manuais = session.get("manuais", {})
    itens = []
    resumo = {"total": 0, "pedir": 0, "sem_saldo": 0, "qtd_pedir": 0}

    if aba_atual:
        itens = montar_aba(aba_atual, catalogo, saldos, manuais)
        resumo["total"] = len(itens)
        for it in itens:
            if it["sem_estoque"]:
                resumo["sem_saldo"] += 1
            if it["pedir"]:
                resumo["pedir"] += 1
                if it["quanto_pedir"]:
                    resumo["qtd_pedir"] += it["quanto_pedir"]

        if filtro == "pedir":
            itens = [i for i in itens if i["pedir"]]
        elif filtro == "sem":
            itens = [i for i in itens if i["sem_estoque"]]

        if q:
            itens = [
                i
                for i in itens
                if q in str(i["codigo"]) or q in (i["descricao"] or "").lower()
            ]

    # resumo global por aba
    contagens = {}
    for a in abas_sorted:
        lista = montar_aba(a["id"], catalogo, saldos, manuais)
        contagens[a["id"]] = sum(1 for i in lista if i["pedir"])

    return render_template(
        "index.html",
        abas=abas_sorted,
        aba_atual=aba_atual,
        itens=itens,
        resumo=resumo,
        filtro=filtro,
        q=q,
        tem_catalogo=bool(abas),
        tem_estoque=bool(saldos),
        estoque_meta=meta,
        contagens=contagens,
        hoje=date.today().strftime("%d/%m/%Y"),
        total_catalogo=sum(a.get("count", 0) for a in abas_sorted),
    )


@app.route("/importar/catalogo", methods=["POST"])
def importar_catalogo():
    f = request.files.get("arquivo")
    password = (request.form.get("senha") or "").strip() or None
    if not f or not f.filename:
        flash("Selecione o arquivo de pedidos (catálogo).", "erro")
        return redirect(url_for("index"))
    try:
        catalog = import_pedido_file(f.read(), f.filename, password)
        if not catalog["abas"]:
            flash("Nenhuma aba de pedido encontrada no arquivo.", "erro")
            return redirect(url_for("index"))
        save_catalogo(catalog)
        total = sum(a["count"] for a in catalog["abas"])
        flash(
            f"Catálogo importado: {len(catalog['abas'])} abas, {total} itens.",
            "ok",
        )
    except Exception as e:
        flash(str(e), "erro")
    return redirect(url_for("index"))


@app.route("/importar/estoque", methods=["POST"])
def importar_estoque():
    f = request.files.get("arquivo")
    if not f or not f.filename:
        flash("Selecione o arquivo EstoqueFarmacia.xlsx.", "erro")
        return redirect(url_for("index"))
    try:
        mapa, meta = parse_estoque_xlsx(f)
        if not mapa:
            flash("Nenhum item de estoque encontrado.", "erro")
            return redirect(url_for("index"))
        save_estoque(mapa, meta)
        flash(f"Estoque importado: {len(mapa)} códigos ({meta.get('arquivo')}).", "ok")
    except Exception as e:
        flash(f"Erro ao ler estoque: {e}", "erro")
    return redirect(url_for("index"))


@app.route("/api/manual", methods=["POST"])
def api_manual():
    data = request.get_json(force=True)
    cod = str(data.get("codigo", ""))
    valor = data.get("valor")
    manuais = session.get("manuais", {})
    if valor in (None, ""):
        manuais.pop(cod, None)
    else:
        manuais[cod] = valor
    session["manuais"] = manuais
    session.modified = True
    return jsonify({"ok": True})


@app.route("/exportar")
def exportar():
    catalogo = load_catalogo()
    saldos = estoque_saldos()
    manuais = session.get("manuais", {})
    so_pedir = request.args.get("so_pedir", "1") == "1"
    aba_filtro = request.args.get("aba")

    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill("solid", fgColor="1B4F72")
    header_font = Font(color="FFFFFF", bold=True)
    sim_fill = PatternFill("solid", fgColor="FDEBD0")
    thin = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    abas = catalogo.get("abas", [])
    for a in abas:
        if aba_filtro and a["id"] != aba_filtro:
            continue
        itens = montar_aba(a["id"], catalogo, saldos, manuais)
        if so_pedir:
            itens = [i for i in itens if i["pedir"] and i["quanto_pedir"]]
        # nome de aba Excel max 31 chars
        safe = re.sub(r"[\\/*?:\[\]]", "-", a["titulo"])[:31] or "Aba"
        ws = wb.create_sheet(safe)
        ws["A1"] = f"PEDIDO MEDICAMENTOS - DATA: {date.today().strftime('%d/%m/%Y')}"
        ws["A2"] = a["titulo"]
        headers = [
            "Cód. AGHU",
            "Medicamento",
            "Est. Mínimo",
            "Ponto de Pedido",
            "Caixa com",
            "Estoque AGHU",
            "Pedir?",
            "Quanto Pedir?",
        ]
        for c, h in enumerate(headers, 1):
            cell = ws.cell(4, c, h)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin

        for r, it in enumerate(itens, 5):
            vals = [
                it["codigo"],
                it["descricao"],
                it["estoque_minimo"],
                it["ponto_pedido"],
                it["caixa_com"],
                "" if it["estoque_aghu"] is None else it["estoque_aghu"],
                "Sim" if it["pedir"] else "Não",
                "" if it["quanto_pedir"] is None else it["quanto_pedir"],
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(r, c, v)
                cell.border = thin
                if it["pedir"]:
                    cell.fill = sim_fill

        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 42
        for col in "CDEFGH":
            ws.column_dimensions[col].width = 14

    if not wb.sheetnames:
        ws = wb.create_sheet("Pedido")
        ws["A1"] = "Nenhum item para pedir."

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"pedido_medicamentos_{date.today().isoformat()}.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/status")
def api_status():
    cat = load_catalogo()
    return jsonify(
        {
            "abas": len(cat.get("abas", [])),
            "itens": sum(a.get("count", 0) for a in cat.get("abas", [])),
            "estoque": len(estoque_saldos()),
            "meta": estoque_meta(),
        }
    )


def bootstrap_estoque_exemplo():
    """Se existir exemplo e ainda não houver estoque, importa."""
    exemplo = DATA / "EstoqueFarmacia-exemplo.xlsx"
    if exemplo.exists() and not ESTOQUE_PATH.exists():
        with exemplo.open("rb") as f:
            class _F:
                filename = exemplo.name
                def read(self_inner):  # noqa
                    return None
            # use path directly
            mapa, meta = parse_estoque_xlsx(exemplo)
            save_estoque(mapa, meta)


if __name__ == "__main__":
    bootstrap_estoque_exemplo()
    print("Pedido de Estoque — http://127.0.0.1:5000")
    app.run(debug=True, port=5000)
