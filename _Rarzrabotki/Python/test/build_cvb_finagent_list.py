# -*- coding: utf-8 -*-
"""Сводный список финагентских ПКО без дочерних ПКО/РКО — из помесячных отчётов ЦВБ.

Источник: reports/cvb_2026-MM.md, разделы «Касса: ручной разбор» (класс C, диагноз
«Фінагент: ЕРП нетто <> Казна»). Ничего не запрашивает у базы — только агрегирует отчёты.
"""
import re
from pathlib import Path

REPORTS = Path(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\reports")

rows = []
for month in range(1, 8):
    f = REPORTS / f"cvb_2026-{month:02d}.md"
    if not f.exists() or f.stat().st_size == 0:
        continue
    for line in f.read_text(encoding="utf-8").splitlines():
        if "Фінагент" not in line or not line.startswith("- "):
            continue
        parts = [p.strip() for p in line[2:].split("|")]
        doc = parts[1] if len(parts) > 1 else ""
        status = parts[2] if len(parts) > 2 else ""
        delta = ""
        m = re.search(r"Δ\s*(-?[\d.,]+)$", line)
        if m:
            delta = m.group(1)
        m_erp = re.search(r"ЕРП нетто\s*([\d\s\u00a0,.]*)<>", status)
        m_kaz = re.search(r"Казна\s*([\d\s\u00a0,.]+)", status)
        rows.append({
            "месяц": f"2026-{month:02d}",
            "документ": doc,
            "ерп": (m_erp.group(1).strip() if m_erp else "").replace("\u00a0", " ") or "—",
            "казна": (m_kaz.group(1).strip() if m_kaz else "").replace("\u00a0", " "),
            "дельта": delta,
        })

out = ["# ЦВБ: финагентские ПКО без дочерних ПКО/РКО (сводно, январь–июль 2026)",
       "",
       "Источник — разделы «Касса: ручной разбор» помесячных отчётов `cvb_2026-MM.md`.",
       "Диагноз движка: «Фінагент: ЕРП нетто <> Казна (перевірте дочірні ПКО/РКО)».",
       "Пустое «ЕРП нетто» = дочерних документов нет ВООБЩЕ (разноска не сделана).",
       ""]

by_month = {}
for r in rows:
    by_month.setdefault(r["месяц"], []).append(r)

out.append("| Месяц | Строк | Из них без дочерних вообще | Σ|Δ| |")
out.append("|---|---|---|---|")
for m in sorted(by_month):
    lst = by_month[m]
    empty = sum(1 for r in lst if r["ерп"] == "—")
    total = 0.0
    for r in lst:
        try:
            total += abs(float(r["дельта"].replace(" ", "").replace(",", ".")))
        except ValueError:
            pass
    out.append(f"| {m} | {len(lst)} | {empty} | {total:,.2f} |")
out.append("")

for m in sorted(by_month):
    out.append(f"## {m} — {len(by_month[m])} документов")
    out.append("")
    out.append("| Документ | ЕРП нетто | Казна | Δ |")
    out.append("|---|---|---|---|")
    for r in sorted(by_month[m], key=lambda x: x["документ"]):
        out.append(f"| {r['документ']} | {r['ерп']} | {r['казна']} | {r['дельта']} |")
    out.append("")

dest = REPORTS / "cvb_finagenty_bez_dochernih.md"
dest.write_text("\n".join(out), encoding="utf-8")
print(f"Всего строк: {len(rows)}")
for m in sorted(by_month):
    lst = by_month[m]
    print(f"  {m}: {len(lst)} (без дочерних вообще: "
          f"{sum(1 for r in lst if r['ерп'] == '—')})")
print(f"Файл: {dest}")
