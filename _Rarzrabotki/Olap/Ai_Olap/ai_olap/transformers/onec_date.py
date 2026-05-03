"""Normalize 1С datetime fields.

In server-clustered 1С (BAS ERP runs on this), `_Date_Time` columns are stored
as plain datetime2 with the actual document date — no 2000-year offset (that
quirk only applies to file-mode bases). pyodbc returns native datetime, so the
transformer mostly truncates to date when requested or to the 1st of the month.

If `to_month=True`, the value becomes the 1st-of-month at midnight (used for
the Period_Month bucket in OlapBASERP fact tables).
"""
from __future__ import annotations

import datetime as dt


def to_month_start(value):
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return dt.datetime(value.year, value.month, 1)
    if isinstance(value, dt.date):
        return dt.date(value.year, value.month, 1)
    return value


def to_date(value):
    if isinstance(value, dt.datetime):
        return value.date()
    return value


def transform(rows: list[dict], *, columns: list[str], mode: str = "month") -> list[dict]:
    """mode='month' -> 1st-of-month; mode='date' -> drop time part; mode='asis' -> noop."""
    fn = {"month": to_month_start, "date": to_date}.get(mode, lambda v: v)
    for row in rows:
        for k in columns:
            if k in row:
                row[k] = fn(row[k])
    return rows
