"""Smoke-test: перевірити що Python COM з'єднується з BaseERP і виконує прості запити.

Коли запускати: після зміни CONN_ERP у config.py, після перезапуску 1С-сервера,
або коли `15_export_to_knowledge_pl.py` починає зависати.

Очікуваний результат: усі 5 кроків виводять "OK" за ~10-30 секунд. Перший
v8.Connect() може займати 5-15 секунд після старту 1С-сервера.

Usage:
    python test_erp_connection.py
"""
import sys
from pathlib import Path

# Скрипт живе в knowledge_PL/Python/PL/test/. Production-скрипти — на рівень вище.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

LOG = Path(__file__).parent / "_test_connection_log.txt"
LOG.write_text("", encoding="utf-8")


def log(msg):
    print(msg, flush=True)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(msg + "\n")
        f.flush()


def main():
    log("=== test_erp_connection.py — smoke-test BaseERP COM ===")

    log("[1/5] import win32com.client ...")
    try:
        import win32com.client  # noqa
        log("  OK (pywin32 installed)")
    except ImportError as e:
        log(f"  FAIL: {e}")
        log("  Install: pip install pywin32")
        sys.exit(1)

    log("[2/5] import _erp_query (uses config.CONN_ERP) ...")
    try:
        from _erp_query import execute_query
        log("  OK")
    except Exception as e:
        log(f"  FAIL: {e}")
        sys.exit(1)

    log("[3/5] ping ERP з найпростішим запитом ...")
    try:
        r = execute_query("ВЫБРАТЬ 1 КАК Ping", max_rows=1)
        assert r and r[0]["Ping"] == 1
        log(f"  OK ({r[0]})")
    except Exception as e:
        log(f"  FAIL: {e}")
        log("  Типові причини: 1С-сервер не запущений, або CONN_ERP у config.py невірний,")
        log("  або поточний користувач не має прав на COM-підключення.")
        sys.exit(1)

    log("[4/5] перевірити доступ до Справочник.А_Статьи_PL ...")
    try:
        r = execute_query(
            "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК N ИЗ Справочник.А_Статьи_PL ГДЕ НЕ ПометкаУдаления",
            max_rows=1,
        )
        n = r[0]["N"]
        log(f"  OK ({n} статей у довіднику)")
        if n < 50:
            log(f"  WARN: очікувалось ~68 статей, знайдено {n}. Перевір чи не пом'ячені.")
    except Exception as e:
        log(f"  FAIL: {e}")
        log("  Довідник А_Статьи_PL відсутній у конфігурації? Перевір через Конфігуратор.")
        sys.exit(1)

    log("[5/5] перевірити Документ.А_ОтчетPL ...")
    try:
        r = execute_query(
            "ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК N ИЗ Документ.А_ОтчетPL ГДЕ НЕ ПометкаУдаления",
            max_rows=1,
        )
        n = r[0]["N"]
        log(f"  OK ({n} документів А_ОтчетPL)")
    except Exception as e:
        log(f"  FAIL: {e}")
        sys.exit(1)

    log("")
    log("=== DONE — COM з'єднання працює ===")
    log("Наступний крок: python ../15_export_to_knowledge_pl.py")


if __name__ == "__main__":
    main()
