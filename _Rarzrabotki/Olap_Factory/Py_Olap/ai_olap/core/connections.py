"""Connection helpers for the three data sources Ai_Olap touches.

- BaseERP SQL backend (pyodbc)   — primary read path for catalogs and registers
- OlapBASERP SQL                  — write target (Stage 2 DDL)
- BaseERP via V83 COM connector   — fallback for virtual tables (.Остатки/.Обороты)
"""
from __future__ import annotations

import os
from typing import Any

import pyodbc
from dotenv import load_dotenv

from .exceptions import ConnectionFailed

load_dotenv()


def _require_env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise ConnectionFailed(f"Environment variable {name} is not set (check .env)")
    return value


def get_baserp_sql() -> pyodbc.Connection:
    """Open a pyodbc connection to the BaseERP MSSQL backend (read-only usage)."""
    dsn = _require_env("BASERP_SQL_DSN")
    try:
        return pyodbc.connect(dsn, timeout=15, readonly=True)
    except Exception as exc:
        raise ConnectionFailed(f"BaseERP SQL connect failed: {exc!r}") from exc


def get_olap_sql() -> pyodbc.Connection:
    """Open a pyodbc connection to OlapBASERP (write target). fast_executemany on cursor."""
    dsn = _require_env("OLAP_SQL_DSN")
    try:
        conn = pyodbc.connect(dsn, timeout=15, autocommit=False)
        conn.cursor().fast_executemany = True
        return conn
    except Exception as exc:
        raise ConnectionFailed(f"OlapBASERP SQL connect failed: {exc!r}") from exc


def get_com_connection() -> Any:
    """Open a COM connection to BaseERP via V83.COMConnector (fallback for virtual tables)."""
    import win32com.client  # imported lazily — heavy dependency, only needed for COM fallback

    conn_str = _require_env("BASERP_COM_CONN")
    try:
        v8 = win32com.client.Dispatch("V83.COMConnector")
        return v8.Connect(conn_str)
    except Exception as exc:
        raise ConnectionFailed(f"BaseERP COM connect failed: {exc!r}") from exc
