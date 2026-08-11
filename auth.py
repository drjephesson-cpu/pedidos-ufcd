# -*- coding: utf-8 -*-
"""Autenticação e usuários (admin / usuario) — Neon ou SQLite."""

from __future__ import annotations

import os
from functools import wraps
from pathlib import Path

from flask import flash, redirect, session, url_for
from sqlalchemy import text
from sqlalchemy.exc import IntegrityError
from werkzeug.security import check_password_hash, generate_password_hash

from db import connect, using_neon

# Admin inicial — senha apenas em hash (não versionar senha em texto puro)
ADMIN_SEED = {
    "username": "jephesson",
    "password_hash": (
        "scrypt:32768:8:1$wzSycIW9DjBbspAJ$"
        "4498fb103ed137aeda0e1f39a4d4c8a0635144d192ab0b809ef99dc5f74af7f7"
        "75a4bd52a41015adbd8ca887da8eb7ce20ba0f8673d011d4284a748916f9602a"
    ),
    "role": "admin",
    "nome": "Jephesson",
}

# Contas conhecidas perdidas no /tmp do Vercel — recriadas no Neon se faltarem.
# Senha temporária: env DEFAULT_USER_PASSWORD ou Alterar@2026
RESTORE_USERS = [
    {
        "username": "lecasser",
        "nome": "LEANDRO CASSER",
        "role": "usuario",
    },
    {
        "username": "camila",
        "nome": "Camila",
        "role": "admin",
    },
]

ROLES = ("admin", "usuario")


class UserStore:
    def __init__(self, db_path: Path | None = None):
        # db_path só usado no SQLite local; no Neon ignora
        self.db_path = db_path
        self._init_db()

    def _sqlite_path(self) -> Path | None:
        if using_neon():
            return None
        return self.db_path

    def _user_exists(self, conn, username: str) -> bool:
        row = conn.execute(
            text(
                "SELECT id FROM users WHERE LOWER(username) = LOWER(:username)"
            ),
            {"username": username},
        ).first()
        return bool(row)

    def _init_db(self) -> None:
        from db import init_db

        init_db(self._sqlite_path())
        with connect(self._sqlite_path()) as conn:
            if not self._user_exists(conn, ADMIN_SEED["username"]):
                conn.execute(
                    text(
                        """
                        INSERT INTO users (username, password_hash, role, nome, ativo)
                        VALUES (:username, :password_hash, :role, :nome, 1)
                        """
                    ),
                    {
                        "username": ADMIN_SEED["username"],
                        "password_hash": ADMIN_SEED["password_hash"],
                        "role": ADMIN_SEED["role"],
                        "nome": ADMIN_SEED["nome"],
                    },
                )

            temp_pw = (os.environ.get("DEFAULT_USER_PASSWORD") or "Alterar@2026").strip()
            for u in RESTORE_USERS:
                if self._user_exists(conn, u["username"]):
                    continue
                conn.execute(
                    text(
                        """
                        INSERT INTO users (username, password_hash, role, nome, ativo)
                        VALUES (:username, :password_hash, :role, :nome, 1)
                        """
                    ),
                    {
                        "username": u["username"],
                        "password_hash": generate_password_hash(temp_pw),
                        "role": u["role"],
                        "nome": u["nome"],
                    },
                )

    def authenticate(self, username: str, password: str) -> dict | None:
        with connect(self._sqlite_path()) as conn:
            row = conn.execute(
                text(
                    """
                    SELECT id, username, password_hash, role, nome, ativo
                    FROM users WHERE LOWER(username) = LOWER(:username)
                    """
                ),
                {"username": username.strip()},
            ).mappings().first()
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
        with connect(self._sqlite_path()) as conn:
            rows = conn.execute(
                text(
                    """
                    SELECT id, username, role, nome, ativo, criado_em
                    FROM users ORDER BY role ASC, username ASC
                    """
                )
            ).mappings().all()
        out = []
        for r in rows:
            d = dict(r)
            d["ativo"] = bool(d.get("ativo"))
            out.append(d)
        return out

    def get_user(self, user_id: int) -> dict | None:
        with connect(self._sqlite_path()) as conn:
            row = conn.execute(
                text(
                    "SELECT id, username, role, nome, ativo FROM users WHERE id = :id"
                ),
                {"id": user_id},
            ).mappings().first()
        if not row:
            return None
        d = dict(row)
        d["ativo"] = bool(d.get("ativo"))
        return d

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
            with connect(self._sqlite_path()) as conn:
                if self._user_exists(conn, username):
                    return False, "Este usuário já existe."
                conn.execute(
                    text(
                        """
                        INSERT INTO users (username, password_hash, role, nome, ativo)
                        VALUES (:username, :password_hash, :role, :nome, 1)
                        """
                    ),
                    {
                        "username": username,
                        "password_hash": generate_password_hash(password),
                        "role": role,
                        "nome": nome,
                    },
                )
        except IntegrityError:
            return False, "Este usuário já existe."
        except Exception as e:
            return False, f"Erro ao salvar usuário: {e}"
        return True, f"Usuário {username} criado e salvo no banco."

    def set_role(self, user_id: int, role: str) -> tuple[bool, str]:
        if role not in ROLES:
            return False, "Papel inválido."
        user = self.get_user(user_id)
        if not user:
            return False, "Usuário não encontrado."
        if user["username"].lower() == ADMIN_SEED["username"].lower() and role != "admin":
            return False, "Não é permitido remover o admin principal."
        with connect(self._sqlite_path()) as conn:
            conn.execute(
                text("UPDATE users SET role = :role WHERE id = :id"),
                {"role": role, "id": user_id},
            )
        return True, "Papel atualizado."

    def set_ativo(self, user_id: int, ativo: bool) -> tuple[bool, str]:
        user = self.get_user(user_id)
        if not user:
            return False, "Usuário não encontrado."
        if user["username"].lower() == ADMIN_SEED["username"].lower() and not ativo:
            return False, "Não é permitido desativar o admin principal."
        with connect(self._sqlite_path()) as conn:
            conn.execute(
                text("UPDATE users SET ativo = :ativo WHERE id = :id"),
                {"ativo": 1 if ativo else 0, "id": user_id},
            )
        return True, "Status atualizado."

    def delete_user(self, user_id: int) -> tuple[bool, str]:
        user = self.get_user(user_id)
        if not user:
            return False, "Usuário não encontrado."
        if user["username"].lower() == ADMIN_SEED["username"].lower():
            return False, "Não é permitido excluir o admin principal."
        with connect(self._sqlite_path()) as conn:
            conn.execute(text("DELETE FROM users WHERE id = :id"), {"id": user_id})
        return True, "Usuário excluído."

    def change_password(self, user_id: int, new_password: str) -> tuple[bool, str]:
        if len(new_password) < 4:
            return False, "Senha deve ter pelo menos 4 caracteres."
        with connect(self._sqlite_path()) as conn:
            res = conn.execute(
                text("UPDATE users SET password_hash = :h WHERE id = :id"),
                {"h": generate_password_hash(new_password), "id": user_id},
            )
            if (res.rowcount or 0) == 0:
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
