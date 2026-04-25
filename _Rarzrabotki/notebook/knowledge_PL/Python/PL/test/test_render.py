"""Перевірити markdown-рендерери на фейкових даних (без ЕРП).

Коли запускати: коли вносиш зміни у функції render_* у _export_pl_knowledge_helpers.py
або у _render_faq.py — щоб пересвідчитися що код рендериться без винятків.

Usage:
    python test_render.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

LOG = Path(__file__).parent / "_test_render_log.txt"
LOG.write_text("", encoding="utf-8")


def log(msg):
    print(msg, flush=True)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(msg + "\n")
        f.flush()


def main():
    log("=== test_render.py — рендерери на фейкових даних ===")

    from _export_pl_knowledge_helpers import (
        render_articles_catalog, render_dds_mapping,
        render_month_dump, render_delta,
        fmt_money, fmt_date, clean_type,
    )

    log("[1/5] fmt_money ...")
    assert fmt_money(0) == "—"
    assert fmt_money(None) == "—"
    assert fmt_money(1234567.89) == "1 234 567.89"
    assert fmt_money(-500) == "−500"
    log("  OK")

    log("[2/5] fmt_date ...")
    assert fmt_date("2026-02-28T21:59:59") == "28.02.2026"
    assert fmt_date("0001-01-01T00:00:00") == ""
    assert fmt_date(None) == ""
    log("  OK")

    log("[3/5] clean_type ...")
    assert clean_type("Документ.РеализацияТоваровУслуг") == "РеализацияТоваровУслуг"
    assert clean_type("Приходный кассовый ордер") == "Приходный кассовый ордер"
    assert clean_type("") == ""
    log("  OK")

    log("[4/5] render_articles_catalog (1 стаття) ...")
    catalog = [{
        "Код": "000000001",
        "Наименование": "Test стаття",
        "ГруппаPLНаим": "Test група",
        "ГруппаPLКод": "000000001",
        "ЭтоГруппа": False,
        "ТипСтатьи": "Расход",
        "ДДСHeaderКод": "УТ-001234",
        "ДДСHeaderНаим": "Test ДДС",
        "СтатьяДоходовНаим": "",
        "СозданоАвтоматически": False,
    }]
    md = render_articles_catalog(catalog)
    assert "Test стаття" in md
    assert "Test ДДС" in md
    log(f"  OK ({len(md)} chars)")

    log("[5/5] render_dds_mapping (1 пара) ...")
    mapping_rows = [{
        "PLКод": "000000001",
        "PLНаим": "Test стаття",
        "PLГруппа": "Test група",
        "PLТип": "Расход",
        "ДДСHeaderКод": "УТ-001234",
        "ДДСHeaderНаим": "Test ДДС",
        "ДДСТЧКод": "УТ-001234",
        "ДДСТЧНаим": "Test ДДС",
    }]
    md = render_dds_mapping(mapping_rows)
    assert "Test стаття" in md
    log(f"  OK ({len(md)} chars)")

    log("")
    log("=== РЕЗУЛЬТАТ: усі рендерери OK ===")


if __name__ == "__main__":
    main()
