# -*- coding: utf-8 -*-
"""Autenticação e usuários (admin / usuario) — Neon ou SQLite."""

from __future__ import annotations

from functools import wraps
from pathlib import Path

from flask import flash, redirect, session, url_for
from sqlalchemy import text
from sqlalchemy.exc import IntegrityError
from werkzeug.security import check_password_hash, generate_password_hash

from db import connect, get_engine, using_neon

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

    def _init_db(self) -> None:
        # tabela users é criada em db.init_db; aqui só garante seed
        from db import init_db

        init_db(self._sqlite_path())
        with connect(self._sqlite_path()) as conn:
            row = conn.execute(
                text(
                    """
                    SELECT id FROM users
                    WHERE LOWER(username) = LOWER(:username)
                    """
                ),
                {"username": ADMIN_SEED["username"]},
            ).first()
            if not row:
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
        return [dict(r) for r in rows]

    def get_user(self, user_id: int) -> dict | None:
        with connect(self._sqlite_path()) as conn:
            row = conn.execute(
                text(
                    "SELECT id, username, role, nome, ativo FROM users WHERE id = :id"
                ),
                {"id": user_id},
            ).mappings().first()
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
            with connect(self._sqlite_path()) as conn:
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
        return True, "Usuário criado."

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
                text(
                    "UPDATE users SET password_hash = :h WHERE id = :id"
                ),
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
