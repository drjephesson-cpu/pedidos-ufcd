# -*- coding: utf-8 -*-
"""Autenticação e usuários (admin / usuario)."""

from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from functools import wraps
from pathlib import Path

from flask import flash, redirect, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash

# Admin inicial — senha apenas em hash (não versionar senha em texto puro)
ADMIN_SEED = {
    "username": "jephesson",
    # hash de senha definida na criação da conta admin
    "password_hash": (
        "scrypt:32768:8:1$wzSycIW9DjBbspAJ$"
        "4498fb103ed137aeda0e1f39a4d4c8a0635144d192ab0b809ef99dc5f74af7f7"
        "75a4bd52a41015adbd8ca887da8eb7ce20ba0f8673d011d4284a748916f9602a"
    ),
    "role": "admin",
    "nome": "Jephesson",
}

ROLES = ("admin", "usuario")


class UserStore:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextmanager
    def _conn(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    def _init_db(self) -> None:
        with self._conn() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL CHECK(role IN ('admin', 'usuario')),
                    nome TEXT NOT NULL DEFAULT '',
                    ativo INTEGER NOT NULL DEFAULT 1,
                    criado_em TEXT NOT NULL DEFAULT (datetime('now','localtime'))
                )
                """
            )
            row = conn.execute(
                "SELECT id FROM users WHERE username = ? COLLATE NOCASE",
                (ADMIN_SEED["username"],),
            ).fetchone()
            if not row:
                conn.execute(
                    """
                    INSERT INTO users (username, password_hash, role, nome)
                    VALUES (?, ?, ?, ?)
                    """,
                    (
                        ADMIN_SEED["username"],
                        ADMIN_SEED["password_hash"],
                        ADMIN_SEED["role"],
                        ADMIN_SEED["nome"],
                    ),
                )

    def authenticate(self, username: str, password: str) -> dict | None:
        with self._conn() as conn:
            row = conn.execute(
                """
                SELECT id, username, password_hash, role, nome, ativo
                FROM users WHERE username = ? COLLATE NOCASE
                """,
                (username.strip(),),
            ).fetchone()
        if not row or not row["ativo"]:
            return None
        if not check_password_hash(row["password_hash"], password):
            return None
        return {
            "id": row["id"],
            "username": row["username"],
            "role": row["role"],
            "nome": row["nome"] or row["username"],
        }

    def list_users(self) -> list[dict]:
        with self._conn() as conn:
            rows = conn.execute(
                """
                SELECT id, username, role, nome, ativo, criado_em
                FROM users ORDER BY role ASC, username ASC
                """
            ).fetchall()
        return [dict(r) for r in rows]

    def get_user(self, user_id: int) -> dict | None:
        with self._conn() as conn:
            row = conn.execute(
                "SELECT id, username, role, nome, ativo FROM users WHERE id = ?",
                (user_id,),
            ).fetchone()
        return dict(row) if row else None

    def create_user(
        self, username: str, password: str, role: str, nome: str = ""
    ) -> tuple[bool, str]:
        username = username.strip()
        nome = (nome or username).strip()
        if not username or not password:
            return False, "Informe usuário e senha."
        if role not in ROLES:
            return False, "Papel inválido."
        if len(password) < 4:
            return False, "Senha deve ter pelo menos 4 caracteres."
        try:
            with self._conn() as conn:
                conn.execute(
                    """
                    INSERT INTO users (username, password_hash, role, nome)
                    VALUES (?, ?, ?, ?)
                    """,
                    (username, generate_password_hash(password), role, nome),
                )
        except sqlite3.IntegrityError:
            return False, "Este usuário já existe."
        return True, "Usuário criado."

    def set_role(self, user_id: int, role: str) -> tuple[bool, str]:
        if role not in ROLES:
            return False, "Papel inválido."
        user = self.get_user(user_id)
        if not user:
            return False, "Usuário não encontrado."
        if user["username"].lower() == ADMIN_SEED["username"].lower() and role != "admin":
            return False, "Não é permitido remover o admin principal."
        with self._conn() as conn:
            conn.execute("UPDATE users SET role = ? WHERE id = ?", (role, user_id))
        return True, "Papel atualizado."

    def set_ativo(self, user_id: int, ativo: bool) -> tuple[bool, str]:
        user = self.get_user(user_id)
        if not user:
            return False, "Usuário não encontrado."
        if user["username"].lower() == ADMIN_SEED["username"].lower() and not ativo:
            return False, "Não é permitido desativar o admin principal."
        with self._conn() as conn:
            conn.execute(
                "UPDATE users SET ativo = ? WHERE id = ?", (1 if ativo else 0, user_id)
            )
        return True, "Status atualizado."

    def delete_user(self, user_id: int) -> tuple[bool, str]:
        user = self.get_user(user_id)
        if not user:
            return False, "Usuário não encontrado."
        if user["username"].lower() == ADMIN_SEED["username"].lower():
            return False, "Não é permitido excluir o admin principal."
        with self._conn() as conn:
            conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
        return True, "Usuário excluído."

    def change_password(self, user_id: int, new_password: str) -> tuple[bool, str]:
        if len(new_password) < 4:
            return False, "Senha deve ter pelo menos 4 caracteres."
        with self._conn() as conn:
            cur = conn.execute(
                "UPDATE users SET password_hash = ? WHERE id = ?",
                (generate_password_hash(new_password), user_id),
            )
            if cur.rowcount == 0:
                return False, "Usuário não encontrado."
        return True, "Senha alterada."


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user"):
            return redirect(url_for("login", next=request_path()))
        return view(*args, **kwargs)

    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        user = session.get("user")
        if not user:
            return redirect(url_for("login", next=request_path()))
        if user.get("role") != "admin":
            flash("Acesso restrito a administradores.", "erro")
            return redirect(url_for("index"))
        return view(*args, **kwargs)

    return wrapped


def request_path() -> str:
    from flask import request

    return request.full_path if request.query_string else request.path


def current_user() -> dict | None:
    return session.get("user")


def is_admin() -> bool:
    user = current_user()
    return bool(user and user.get("role") == "admin")
