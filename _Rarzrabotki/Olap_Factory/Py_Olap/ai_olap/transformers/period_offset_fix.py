"""Знімає +2000 рік-офсет із BaseERP `_Date_Time` колонки, обрізає час і деривує Period_Month.

BaseERP cluster backend зберігає всі datetime поля з +2000 років (документ за
2026-02-15 у MSSQL — `4026-02-15`). Цей transformer:

1. Якщо `Period` це datetime з роком ≥ `epoch_threshold` (default 3000) — віднімає
   `offset_years` (default 2000), повертаючи Period у "людському" році (2026).
2. Якщо `strip_time=True` (default) і Period — datetime, обрізає компонент часу
   до 00:00:00.000 (треба щоб PowerBI Date Range slicer на DATETIME колонці працював).
3. Якщо `derive_period_month=True` (default) — виставляє `Period_Month` як
   `date(Period.year, Period.month, 1)`.

Якщо Period = None або вже < epoch_threshold (тобто без офсету) — strip_time все
одно застосовується якщо там datetime.
"""
from __future__ import annotations

import datetime as dt


def _strip_offset(value, offset_years: int, threshold: int):
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        if value.year >= threshold:
            return value.replace(year=value.year - offset_years)
        return value
    if isinstance(value, dt.date):
        if value.year >= threshold:
            return value.replace(year=value.year - offset_years)
        return value
    return value


def transform(
    rows: list[dict],
    *,
    source_column: str = "Period",
    period_month_column: str = "Period_Month",
    offset_years: int = 2000,
    epoch_threshold: int = 3000,
    derive_period_month: bool = True,
    strip_time: bool = True,
) -> list[dict]:
    for row in rows:
        v = row.get(source_column)
        v_fixed = _strip_offset(v, offset_years, epoch_threshold)
        if strip_time and isinstance(v_fixed, dt.datetime):
            v_fixed = v_fixed.replace(hour=0, minute=0, second=0, microsecond=0)
        row[source_column] = v_fixed
        if derive_period_month and isinstance(v_fixed, (dt.date, dt.datetime)):
            row[period_month_column] = dt.date(v_fixed.year, v_fixed.month, 1)
        elif derive_period_month and v_fixed is None:
            row[period_month_column] = None
    return rows
