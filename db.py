# -*- coding: utf-8 -*-
"""Persistência de pedidos — Neon (Postgres) ou SQLite local."""

from __future__ import annotations

import os
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from typing import Any

from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

_engine: Engine | None = None


def _normalize_db_url(url: str) -> str:
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://") :]
    # SQLAlchemy + psycopg3 (pacote `psycopg`) precisa do dialeto explícito
    if url.startswith("postgresql://") and "+psycopg" not in url.split("://", 1)[0]:
        url = "postgresql+psycopg://" + url[len("postgresql://") :]
    elif url.startswith("postgresql+psycopg2://"):
        url = "postgresql+psycopg://" + url[len("postgresql+psycopg2://") :]
    return url


def get_database_url(sqlite_path: Path | None = None) -> str:
    url = (os.environ.get("DATABASE_URL") or "").strip()
    if url:
        return _normalize_db_url(url)
    path = sqlite_path or Path("data/pedidos.db")
    path.parent.mkdir(parents=True, exist_ok=True)
    return f"sqlite:///{path.as_posix()}"


def get_engine(sqlite_path: Path | None = None) -> Engine:
    global _engine
    if _engine is None:
        url = get_database_url(sqlite_path)
        connect_args = {}
        if url.startswith("sqlite"):
            connect_args["check_same_thread"] = False
        _engine = create_engine(url, pool_pre_ping=True, connect_args=connect_args)
    return _engine


def using_neon() -> bool:
    return bool((os.environ.get("DATABASE_URL") or "").strip())


def reset_engine() -> None:
    global _engine
    if _engine is not None:
        _engine.dispose()
    _engine = None


@contextmanager
def connect(sqlite_path: Path | None = None):
    eng = get_engine(sqlite_path)
    with eng.begin() as conn:
        yield conn


def init_db(sqlite_path: Path | None = None) -> None:
    dialect = get_engine(sqlite_path).dialect.name
    id_type = (
        "SERIAL PRIMARY KEY"
        if dialect == "postgresql"
        else "INTEGER PRIMARY KEY AUTOINCREMENT"
    )
    ts_default = "NOW()" if dialect == "postgresql" else "CURRENT_TIMESTAMP"

    create_stmts = [
        f"""
        CREATE TABLE IF NOT EXISTS pedidos (
            id {id_type},
            data_pedido DATE NOT NULL,
            usuario TEXT,
            observacao TEXT,
            unidade TEXT,
            criado_em TIMESTAMP NOT NULL DEFAULT {ts_default}
        )
        """,
        f"""
        CREATE TABLE IF NOT EXISTS pedido_itens (
            id {id_type},
            pedido_id INTEGER NOT NULL,
            codigo BIGINT,
            descricao TEXT NOT NULL,
            aba TEXT,
            quantidade REAL NOT NULL,
            origem TEXT NOT NULL DEFAULT 'auto',
            estoque_aghu REAL,
            estoque_minimo REAL,
            ponto_pedido REAL,
            FOREIGN KEY (pedido_id) REFERENCES pedidos(id) ON DELETE CASCADE
        )
        """,
        f"""
        CREATE TABLE IF NOT EXISTS users (
            id {id_type},
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL,
            nome TEXT NOT NULL DEFAULT '',
            ativo INTEGER NOT NULL DEFAULT 1,
            criado_em TIMESTAMP NOT NULL DEFAULT {ts_default}
        )
        """,
        f"""
        CREATE TABLE IF NOT EXISTS estoque_unidade (
            unidade TEXT PRIMARY KEY,
            saldos_json TEXT NOT NULL,
            meta_json TEXT,
            atualizado_em TIMESTAMP NOT NULL DEFAULT {ts_default}
        )
        """,
    ]
    index_stmts = [
        "CREATE INDEX IF NOT EXISTS idx_pedidos_data ON pedidos (data_pedido DESC)",
        "CREATE INDEX IF NOT EXISTS idx_pedidos_unidade ON pedidos (unidade)",
        "CREATE INDEX IF NOT EXISTS idx_pedidos_usuario ON pedidos (usuario)",
        "CREATE INDEX IF NOT EXISTS idx_itens_aba ON pedido_itens (aba)",
    ]
    with connect(sqlite_path) as conn:
        for s in create_stmts:
            conn.execute(text(s))
        # migração ANTES dos índices que usam a coluna
        if dialect == "postgresql":
            exists = conn.execute(
                text(
                    """
                    SELECT 1 FROM information_schema.columns
                    WHERE table_schema = 'public'
                      AND table_name = 'pedidos'
                      AND column_name = 'unidade'
                    """
                )
            ).first()
            if not exists:
                conn.execute(text("ALTER TABLE pedidos ADD COLUMN unidade TEXT"))
        else:
            names = {
                r[1]
                for r in conn.execute(text("PRAGMA table_info(pedidos)")).fetchall()
            }
            if "unidade" not in names:
                conn.execute(text("ALTER TABLE pedidos ADD COLUMN unidade TEXT"))

        # backfill simples a partir da observação
        conn.execute(
            text(
                """
                UPDATE pedidos
                SET unidade = 'cc'
                WHERE (unidade IS NULL OR TRIM(unidade) = '')
                  AND observacao ILIKE '%Centro Cir_rgico%'
                """
                if dialect == "postgresql"
                else """
                UPDATE pedidos
                SET unidade = 'cc'
                WHERE (unidade IS NULL OR TRIM(unidade) = '')
                  AND observacao LIKE '%Centro Cir_rgico%'
                """
            )
        )
        conn.execute(
            text(
                """
                UPDATE pedidos
                SET unidade = 'ufcd'
                WHERE (unidade IS NULL OR TRIM(unidade) = '')
                  AND observacao ILIKE '%UFCD%'
                """
                if dialect == "postgresql"
                else """
                UPDATE pedidos
                SET unidade = 'ufcd'
                WHERE (unidade IS NULL OR TRIM(unidade) = '')
                  AND observacao LIKE '%UFCD%'
                """
            )
        )
        conn.execute(
            text(
                """
                UPDATE pedidos
                SET unidade = 'ufcd'
                WHERE unidade IS NULL OR TRIM(unidade) = ''
                """
            )
        )

        for s in index_stmts:
            conn.execute(text(s))


def _infer_unidade(unidade: str | None, observacao: str | None) -> str:
    u = (unidade or "").strip().lower()
    if u in ("cc", "centro", "centro_cirurgico", "bloco"):
        return "cc"
    if u in ("ufcd",):
        return "ufcd"
    obs = observacao or ""
    if "Centro Cirúrgico" in obs or "[cc]" in obs.lower():
        return "cc"
    if "UFCD" in obs:
        return "ufcd"
    return u or "ufcd"


def criar_pedido(
    data_pedido: date,
    usuario: str | None,
    observacao: str | None,
    itens: list[dict[str, Any]],
    sqlite_path: Path | None = None,
    unidade: str | None = None,
) -> int:
    if not itens:
        raise ValueError("Pedido sem itens.")

    unidade_id = _infer_unidade(unidade or (itens[0].get("unidade") if itens else None), observacao)
    is_pg = get_engine(sqlite_path).dialect.name == "postgresql"
    with connect(sqlite_path) as conn:
        if is_pg:
            pedido_id = conn.execute(
                text(
                    """
                    INSERT INTO pedidos (data_pedido, usuario, observacao, unidade)
                    VALUES (:data_pedido, :usuario, :observacao, :unidade)
                    RETURNING id
                    """
                ),
                {
                    "data_pedido": data_pedido.isoformat(),
                    "usuario": usuario or "",
                    "observacao": observacao or "",
                    "unidade": unidade_id,
                },
            ).scalar_one()
        else:
            res = conn.execute(
                text(
                    """
                    INSERT INTO pedidos (data_pedido, usuario, observacao, unidade)
                    VALUES (:data_pedido, :usuario, :observacao, :unidade)
                    """
                ),
                {
                    "data_pedido": data_pedido.isoformat(),
                    "usuario": usuario or "",
                    "observacao": observacao or "",
                    "unidade": unidade_id,
                },
            )
            pedido_id = int(res.lastrowid)

        for it in itens:
            conn.execute(
                text(
                    """
                    INSERT INTO pedido_itens (
                        pedido_id, codigo, descricao, aba, quantidade, origem,
                        estoque_aghu, estoque_minimo, ponto_pedido
                    ) VALUES (
                        :pedido_id, :codigo, :descricao, :aba, :quantidade, :origem,
                        :estoque_aghu, :estoque_minimo, :ponto_pedido
                    )
                    """
                ),
                {
                    "pedido_id": pedido_id,
                    "codigo": it.get("codigo"),
                    "descricao": it.get("descricao") or "",
                    "aba": it.get("aba") or "",
                    "quantidade": float(it.get("quantidade") or 0),
                    "origem": it.get("origem") or "auto",
                    "estoque_aghu": it.get("estoque_aghu"),
                    "estoque_minimo": it.get("estoque_minimo"),
                    "ponto_pedido": it.get("ponto_pedido"),
                },
            )
    return int(pedido_id)


def listar_filtros_historico(sqlite_path: Path | None = None) -> dict:
    with connect(sqlite_path) as conn:
        usuarios = [
            r[0]
            for r in conn.execute(
                text(
                    """
                    SELECT DISTINCT usuario FROM pedidos
                    WHERE usuario IS NOT NULL AND TRIM(usuario) <> ''
                    ORDER BY usuario
                    """
                )
            ).fetchall()
        ]
        abas = [
            r[0]
            for r in conn.execute(
                text(
                    """
                    SELECT DISTINCT aba FROM pedido_itens
                    WHERE aba IS NOT NULL AND TRIM(aba) <> ''
                    ORDER BY aba
                    """
                )
            ).fetchall()
        ]
        unidades = [
            r[0]
            for r in conn.execute(
                text(
                    """
                    SELECT DISTINCT unidade FROM pedidos
                    WHERE unidade IS NOT NULL AND TRIM(unidade) <> ''
                    ORDER BY unidade
                    """
                )
            ).fetchall()
        ]
    return {"usuarios": usuarios, "abas": abas, "unidades": unidades}


def listar_pedidos(
    sqlite_path: Path | None = None,
    *,
    unidade: str | None = None,
    aba: str | None = None,
    usuario: str | None = None,
) -> list[dict]:
    is_pg = get_engine(sqlite_path).dialect.name == "postgresql"
    abas_expr = (
        "STRING_AGG(DISTINCT NULLIF(TRIM(i.aba), ''), ', ')"
        if is_pg
        else "GROUP_CONCAT(DISTINCT NULLIF(TRIM(i.aba), ''))"
    )
    where = ["1=1"]
    params: dict[str, Any] = {}
    if unidade:
        where.append("COALESCE(p.unidade, '') = :unidade")
        params["unidade"] = unidade
    if usuario:
        where.append("COALESCE(p.usuario, '') = :usuario")
        params["usuario"] = usuario
    if aba:
        where.append(
            "EXISTS (SELECT 1 FROM pedido_itens ix WHERE ix.pedido_id = p.id AND ix.aba = :aba)"
        )
        params["aba"] = aba

    sql = f"""
        SELECT p.id, p.data_pedido, p.usuario, p.observacao, p.unidade, p.criado_em,
               COUNT(i.id) AS qtd_itens,
               COALESCE(SUM(i.quantidade), 0) AS qtd_total,
               {abas_expr} AS abas
        FROM pedidos p
        LEFT JOIN pedido_itens i ON i.pedido_id = p.id
        WHERE {' AND '.join(where)}
        GROUP BY p.id, p.data_pedido, p.usuario, p.observacao, p.unidade, p.criado_em
        ORDER BY p.unidade ASC NULLS LAST, p.usuario ASC NULLS LAST,
                 p.data_pedido DESC, p.id DESC
    """
    # SQLite não tem NULLS LAST — ordem simples
    if not is_pg:
        sql = f"""
        SELECT p.id, p.data_pedido, p.usuario, p.observacao, p.unidade, p.criado_em,
               COUNT(i.id) AS qtd_itens,
               COALESCE(SUM(i.quantidade), 0) AS qtd_total,
               {abas_expr} AS abas
        FROM pedidos p
        LEFT JOIN pedido_itens i ON i.pedido_id = p.id
        WHERE {' AND '.join(where)}
        GROUP BY p.id, p.data_pedido, p.usuario, p.observacao, p.unidade, p.criado_em
        ORDER BY COALESCE(p.unidade, ''), COALESCE(p.usuario, ''),
                 p.data_pedido DESC, p.id DESC
        """

    with connect(sqlite_path) as conn:
        rows = conn.execute(text(sql), params).mappings().all()

    out = []
    for r in rows:
        d = dict(r)
        d["unidade"] = _infer_unidade(d.get("unidade"), d.get("observacao"))
        abas_raw = d.get("abas") or ""
        if isinstance(abas_raw, str):
            d["abas_lista"] = [a.strip() for a in abas_raw.split(",") if a.strip()]
        else:
            d["abas_lista"] = []
        out.append(d)
    return out


def obter_pedido(pedido_id: int, sqlite_path: Path | None = None) -> dict | None:
    with connect(sqlite_path) as conn:
        ped = conn.execute(
            text(
                """
                SELECT id, data_pedido, usuario, observacao, unidade, criado_em
                FROM pedidos WHERE id = :id
                """
            ),
            {"id": pedido_id},
        ).mappings().first()
        if not ped:
            return None
        itens = conn.execute(
            text(
                """
                SELECT id, codigo, descricao, aba, quantidade, origem,
                       estoque_aghu, estoque_minimo, ponto_pedido
                FROM pedido_itens
                WHERE pedido_id = :id
                ORDER BY aba, descricao
                """
            ),
            {"id": pedido_id},
        ).mappings().all()
    out = dict(ped)
    out["unidade"] = _infer_unidade(out.get("unidade"), out.get("observacao"))
    out["itens"] = [dict(i) for i in itens]
    return out


def excluir_pedido(pedido_id: int, sqlite_path: Path | None = None) -> bool:
    with connect(sqlite_path) as conn:
        res = conn.execute(
            text("DELETE FROM pedidos WHERE id = :id"), {"id": pedido_id}
        )
        return (res.rowcount or 0) > 0


def load_estoque_db(unidade: str, sqlite_path: Path | None = None) -> dict | None:
    """Retorna {saldos, meta} do banco, ou None se não houver."""
    import json

    with connect(sqlite_path) as conn:
        row = conn.execute(
            text(
                """
                SELECT saldos_json, meta_json FROM estoque_unidade
                WHERE unidade = :unidade
                """
            ),
            {"unidade": unidade},
        ).mappings().first()
    if not row:
        return None
    try:
        saldos = json.loads(row["saldos_json"] or "{}")
    except Exception:
        saldos = {}
    try:
        meta = json.loads(row["meta_json"] or "{}")
    except Exception:
        meta = {}
    return {"saldos": saldos, "meta": meta}


def save_estoque_db(
    unidade: str,
    saldos: dict,
    meta: dict | None = None,
    sqlite_path: Path | None = None,
) -> None:
    import json

    saldos_json = json.dumps(saldos, ensure_ascii=False)
    meta_json = json.dumps(meta or {}, ensure_ascii=False)
    is_pg = get_engine(sqlite_path).dialect.name == "postgresql"
    with connect(sqlite_path) as conn:
        if is_pg:
            conn.execute(
                text(
                    """
                    INSERT INTO estoque_unidade (unidade, saldos_json, meta_json, atualizado_em)
                    VALUES (:unidade, :saldos_json, :meta_json, NOW())
                    ON CONFLICT (unidade) DO UPDATE SET
                        saldos_json = EXCLUDED.saldos_json,
                        meta_json = EXCLUDED.meta_json,
                        atualizado_em = NOW()
                    """
                ),
                {
                    "unidade": unidade,
                    "saldos_json": saldos_json,
                    "meta_json": meta_json,
                },
            )
        else:
            conn.execute(
                text(
                    """
                    INSERT INTO estoque_unidade (unidade, saldos_json, meta_json, atualizado_em)
                    VALUES (:unidade, :saldos_json, :meta_json, CURRENT_TIMESTAMP)
                    ON CONFLICT(unidade) DO UPDATE SET
                        saldos_json = excluded.saldos_json,
                        meta_json = excluded.meta_json,
                        atualizado_em = CURRENT_TIMESTAMP
                    """
                ),
                {
                    "unidade": unidade,
                    "saldos_json": saldos_json,
                    "meta_json": meta_json,
                },
            )
