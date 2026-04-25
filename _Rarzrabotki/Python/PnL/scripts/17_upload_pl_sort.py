"""
Заповнює реквізит `Сорт` у каталогах А_ГруппаСтатей_PL і А_Статьи_PL
з JSON-файла, згенерованого скриптом 16_extract_pl_sort_order.py.

Поведінка:
- Пошук існуючого елемента через НайтиПоНаименованию (точний пошук).
- Якщо знайдено — оновити Сорт, записати з ОбменДанными.Загрузка=True (обхід ПередЗаписью гардів).
- А_РучнаяКорректировка=Истина НЕ блокує оновлення Сорт — це організаційне
  поле відображення, не маппінг ДДС (рішення користувача 2026-04-25).
- Якщо НЕ знайдено — записати в лог, не падати.

Запуск (інтерактивний — пише у БД):
    C:\\Python313\\python.exe scripts/17_upload_pl_sort.py
    C:\\Python313\\python.exe scripts/17_upload_pl_sort.py --dry-run   # тільки лог, без запису
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import config  # noqa: E402
from utils.com_connect import connect_erp  # noqa: E402

JSON_PATH = config.JSON_DIR / "16_pl_sort_order.json"


def upload_groups(conn, groups: list[dict], dry_run: bool):
    mgr = conn.Справочники.А_ГруппаСтатей_PL
    stats = {"updated": 0, "not_found": 0, "unchanged": 0}
    for g in groups:
        ref = mgr.НайтиПоНаименованию(g["name"], True)  # exact match
        if ref is None or ref.Пустая():
            print(f"  [GROUP][NOT FOUND] {g['name']}")
            stats["not_found"] += 1
            continue
        obj = ref.ПолучитьОбъект()
        old = float(obj.Сорт or 0)
        new = float(g["sort"])
        if abs(old - new) < 0.5:
            stats["unchanged"] += 1
            continue
        if dry_run:
            print(f"  [GROUP][DRY] {g['name']:<60} {old:>5.0f} → {new:>5.0f}")
        else:
            obj.Сорт = new
            obj.ОбменДанными.Загрузка = True
            obj.Записать()
            print(f"  [GROUP][UPDATED] {g['name']:<60} {old:>5.0f} → {new:>5.0f}")
        stats["updated"] += 1
    return stats


def upload_articles(conn, articles: list[dict], dry_run: bool):
    mgr = conn.Справочники.А_Статьи_PL
    stats = {"updated": 0, "not_found": 0, "unchanged": 0}
    for a in articles:
        ref = mgr.НайтиПоНаименованию(a["name"], True)
        if ref is None or ref.Пустая():
            print(f"  [ART][NOT FOUND] [{a['group'][:30]}] {a['name']}")
            stats["not_found"] += 1
            continue
        obj = ref.ПолучитьОбъект()
        # Сорт — організаційне поле, оновлюємо навіть для locked (А_РучнаяКорректировка=Истина).
        old = float(obj.Сорт or 0)
        new = float(a["sort"])
        if abs(old - new) < 0.5:
            stats["unchanged"] += 1
            continue
        if dry_run:
            print(f"  [ART][DRY] [{a['group'][:30]:<30}] {a['name']:<55} {old:>5.0f} → {new:>5.0f}")
        else:
            obj.Сорт = new
            obj.ОбменДанными.Загрузка = True
            obj.Записать()
            print(f"  [ART][UPD] [{a['group'][:30]:<30}] {a['name']:<55} {old:>5.0f} → {new:>5.0f}")
        stats["updated"] += 1
    return stats


def main():
    dry_run = "--dry-run" in sys.argv
    if not JSON_PATH.exists():
        print(f"FAIL: {JSON_PATH} не існує. Запустіть спочатку 16_extract_pl_sort_order.py")
        sys.exit(1)
    payload = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    groups = payload["groups"]
    articles = payload["articles"]

    print(f"Завантажуємо Сорт із {JSON_PATH}")
    print(f"  groups: {len(groups)}  articles: {len(articles)}")
    print(f"  mode: {'DRY-RUN' if dry_run else 'WRITE'}")
    print()

    conn = connect_erp()

    print("=== ГРУПИ ===")
    g_stats = upload_groups(conn, groups, dry_run)
    print(f"  Підсумок груп: {g_stats}")

    print()
    print("=== СТАТТІ ===")
    a_stats = upload_articles(conn, articles, dry_run)
    print(f"  Підсумок статей: {a_stats}")

    print()
    print("DONE" + (" (DRY-RUN, нічого не записано)" if dry_run else ""))


if __name__ == "__main__":
    main()
