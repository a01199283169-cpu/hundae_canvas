"""SQLite / Supabase(Postgres) 공통 DB 연결."""

from __future__ import annotations

import os
import re
import sqlite3
from pathlib import Path
from typing import Any

from src.config_loader import ROOT_DIR, load_config, resolve_path
from src.settings import get_database_path


class Row(dict):
    """dict + 정수 인덱스 접근 (fetchone()[0] 호환)."""

    def __getitem__(self, key: int | str) -> Any:
        if isinstance(key, int):
            return list(self.values())[key]
        return super().__getitem__(key)


def is_postgres() -> bool:
    url = os.getenv("DATABASE_URL", "").strip()
    return url.startswith("postgres://") or url.startswith("postgresql://")


def _sqlite_path() -> Path:
    return get_database_path(load_config())


def _adapt_sql(sql: str) -> str:
    """SQLite SQL → Postgres 변환 (? 플레이스홀더, 함수 차이)."""
    if not is_postgres():
        return sql

    s = sql
    # GLOB '????-??-??' → 정규식
    s = re.sub(
        r"(\S+)\s+GLOB\s+'\?\?\?\?-\?\?-\?\?'",
        r"\1 ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'",
        s,
        flags=re.IGNORECASE,
    )
    # GROUP_CONCAT(a, 'sep') → STRING_AGG
    s = re.sub(
        r"GROUP_CONCAT\((.+?),\s*'([^']*)'\)",
        r"STRING_AGG(\1, '\2')",
        s,
        flags=re.IGNORECASE | re.DOTALL,
    )
    s = s.replace("?", "%s")
    return s


def has_image_clause(alias: str = "") -> str:
    """has_image 참 — SQLite(0/1) / Postgres(boolean)."""
    col = f"{alias}has_image" if alias else "has_image"
    return f"{col} IS TRUE" if is_postgres() else f"{col}=1"


def date_like_expr(expr: str) -> str:
    """월별 LIKE — Postgres DATE는 ::text 캐스트."""
    return f"{expr}::text LIKE %s" if is_postgres() else f"{expr} LIKE ?"


def order_no_order() -> str:
    return "CAST(order_no AS INTEGER)"


class CursorWrapper:
    def __init__(self, cursor: Any, is_pg: bool) -> None:
        self._cur = cursor
        self._is_pg = is_pg
        self._last_id: int | None = None

    @property
    def lastrowid(self) -> int | None:
        if self._last_id is not None:
            return self._last_id
        return getattr(self._cur, "lastrowid", None)

    @property
    def rowcount(self) -> int:
        return self._cur.rowcount

    def fetchone(self) -> Row | None:
        row = self._cur.fetchone()
        if row is None:
            return None
        if isinstance(row, dict):
            return Row(row)
        if hasattr(row, "keys"):
            return Row(dict(row))
        desc = self._cur.description or []
        cols = [d[0] for d in desc]
        return Row(dict(zip(cols, row)))

    def fetchall(self) -> list[Row]:
        out: list[Row] = []
        for raw in self._cur.fetchall():
            if isinstance(raw, dict):
                out.append(Row(raw))
            elif hasattr(raw, "keys"):
                out.append(Row(dict(raw)))
            else:
                desc = self._cur.description or []
                cols = [d[0] for d in desc]
                out.append(Row(dict(zip(cols, raw))))
        return out


class ConnectionWrapper:
    """SQLite / Postgres 통합 연결."""

    def __init__(self, conn: Any, is_pg: bool) -> None:
        self._conn = conn
        self._is_pg = is_pg

    def execute(self, sql: str, params: tuple | list = ()) -> CursorWrapper:
        sql = _adapt_sql(sql)
        cur = self._conn.cursor()
        cur.execute(sql, tuple(params))
        return CursorWrapper(cur, self._is_pg)

    def executescript(self, sql: str) -> None:
        if self._is_pg:
            for stmt in sql.split(";"):
                stmt = stmt.strip()
                if stmt:
                    self.execute(stmt)
        else:
            self._conn.executescript(sql)

    def commit(self) -> None:
        self._conn.commit()

    def close(self) -> None:
        self._conn.close()


def connect() -> ConnectionWrapper:
    """DB 연결 — DATABASE_URL(Postgres) 또는 로컬 SQLite."""
    if is_postgres():
        import psycopg2
        from psycopg2.extras import RealDictCursor

        url = os.environ["DATABASE_URL"].strip()
        conn = psycopg2.connect(url, cursor_factory=RealDictCursor)
        return ConnectionWrapper(conn, True)

    db_path = _sqlite_path()
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return ConnectionWrapper(conn, False)


def init_postgres_schema(conn: ConnectionWrapper) -> None:
    """Supabase 마이그레이션 SQL 실행."""
    mig = ROOT_DIR / "supabase" / "migrations" / "001_initial.sql"
    sql = mig.read_text(encoding="utf-8")
    conn.executescript(sql)
    conn.commit()


def insert_returning_id(conn: ConnectionWrapper, sql: str, params: tuple | list) -> int:
    """INSERT 후 id 반환."""
    if is_postgres():
        sql = _adapt_sql(sql.rstrip().rstrip(";") + " RETURNING id")
        cur = conn.execute(sql, params)
        row = cur.fetchone()
        conn.commit()
        return int(row["id"] if row else 0)
    cur = conn.execute(sql, params)
    conn.commit()
    return int(cur.lastrowid or 0)
