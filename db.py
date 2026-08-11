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
        return "postgresql://" + url[len("postgres://") :]
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

    stmts = [
        f"""
        CREATE TABLE IF NOT EXISTS pedidos (
            id {id_type},
            data_pedido DATE NOT NULL,
            usuario TEXT,
            observacao TEXT,
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
        "CREATE INDEX IF NOT EXISTS idx_pedidos_data ON pedidos (data_pedido DESC)",
    ]
    with connect(sqlite_path) as conn:
        for s in stmts:
            conn.execute(text(s))


def criar_pedido(
    data_pedido: date,
    usuario: str | None,
    observacao: str | None,
    itens: list[dict[str, Any]],
    sqlite_path: Path | None = None,
) -> int:
    if not itens:
        raise ValueError("Pedido sem itens.")

    is_pg = get_engine(sqlite_path).dialect.name == "postgresql"
    with connect(sqlite_path) as conn:
        if is_pg:
            pedido_id = conn.execute(
                text(
                    """
                    INSERT INTO pedidos (data_pedido, usuario, observacao)
                    VALUES (:data_pedido, :usuario, :observacao)
                    RETURNING id
                    """
                ),
                {
                    "data_pedido": data_pedido.isoformat(),
                    "usuario": usuario or "",
                    "observacao": observacao or "",
                },
            ).scalar_one()
        else:
            res = conn.execute(
                text(
                    """
                    INSERT INTO pedidos (data_pedido, usuario, observacao)
                    VALUES (:data_pedido, :usuario, :observacao)
                    """
                ),
                {
                    "data_pedido": data_pedido.isoformat(),
                    "usuario": usuario or "",
                    "observacao": observacao or "",
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


def listar_pedidos(sqlite_path: Path | None = None) -> list[dict]:
    with connect(sqlite_path) as conn:
        rows = conn.execute(
            text(
                """
                SELECT p.id, p.data_pedido, p.usuario, p.observacao, p.criado_em,
                       COUNT(i.id) AS qtd_itens,
                       COALESCE(SUM(i.quantidade), 0) AS qtd_total
                FROM pedidos p
                LEFT JOIN pedido_itens i ON i.pedido_id = p.id
                GROUP BY p.id, p.data_pedido, p.usuario, p.observacao, p.criado_em
                ORDER BY p.data_pedido DESC, p.id DESC
                """
            )
        ).mappings().all()
    return [dict(r) for r in rows]


def obter_pedido(pedido_id: int, sqlite_path: Path | None = None) -> dict | None:
    with connect(sqlite_path) as conn:
        ped = conn.execute(
            text(
                """
                SELECT id, data_pedido, usuario, observacao, criado_em
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
    out["itens"] = [dict(i) for i in itens]
    return out


def excluir_pedido(pedido_id: int, sqlite_path: Path | None = None) -> bool:
    with connect(sqlite_path) as conn:
        res = conn.execute(
            text("DELETE FROM pedidos WHERE id = :id"), {"id": pedido_id}
        )
        return (res.rowcount or 0) > 0
