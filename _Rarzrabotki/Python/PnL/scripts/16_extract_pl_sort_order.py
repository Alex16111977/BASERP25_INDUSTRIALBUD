"""
Витягає канонічний порядок груп і статей PL з Excel-файла "Лютий 2026"
по списку листів з документа `А_РасшифровкаЛистов №000000003 від 28.02.2026`.

Алгоритм:
1. COM-запит до 1С → отримати список листів з ТЧ Расшифровка документа.
2. Парсити Excel через excel_parser.parse_workbook(), фільтрувати тільки ці листи.
3. Для кожної групи зібрати позиції на кожному листі (індекс рядка серед всіх груп листа).
4. Median position → ранг → Сорт = (rank+1) * 100.
5. Для статей всередині групи: позиція в межах групи на кожному листі → median → ранг → *100.
6. Зберегти у data/json/16_pl_sort_order.json.

Запуск:
    C:\\Python313\\python.exe scripts/16_extract_pl_sort_order.py
"""
import datetime
import json
import sys
from collections import defaultdict
from pathlib import Path
from statistics import median

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import config  # noqa: E402
from utils.com_connect import connect_erp  # noqa: E402
from utils.excel_parser import parse_workbook  # noqa: E402

DOC_NUMBER = "000000003"
DOC_DATE = datetime.datetime(2026, 2, 28, 12, 0, 0)
EXCEL_LABEL = "Лютий 2026"
OUT_JSON = config.JSON_DIR / "16_pl_sort_order.json"


def get_used_sheets(conn) -> list[str]:
    """Читає список ИмяЛиста з ТЧ Расшифровка документа А_РасшифровкаЛистов №000000003."""
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ Т.ИмяЛиста КАК ИмяЛиста, Т.НомерСтроки КАК НомерСтроки\n"
        "ИЗ Документ.А_РасшифровкаЛистов.Расшифровка КАК Т\n"
        'ГДЕ Т.Ссылка.Номер = "' + DOC_NUMBER + '"\n'
        "УПОРЯДОЧИТЬ ПО Т.НомерСтроки"
    )
    sel = q.Выполнить().Выбрать()
    sheets = []
    while sel.Следующий():
        s = str(sel.ИмяЛиста or "").strip()
        if s:
            sheets.append(s)
    return sheets


def load_excel_sheets(xlsx_path: str, month_header: str, allowed_sheets: set):
    """Парсить тільки ті листи з Excel що в allowed_sheets. Повертає dict {sheet_name: rows}.
    rows — список dict з полями row, article, group (як з parse_workbook).
    """
    by_sheet = {}
    for sheet_data in parse_workbook(xlsx_path, month_header):
        name = sheet_data["sheet_name"].strip()
        if name not in allowed_sheets:
            continue
        by_sheet[name] = sheet_data["rows"]
    return by_sheet


def aggregate_group_order(by_sheet: dict) -> list[str]:
    """Для кожної групи збирає позиції її першої появи серед груп на листі.
    Повертає список груп у порядку зростання median-позиції.
    """
    positions = defaultdict(list)  # group_name → [pos1, pos2, ...]

    for sheet_name, rows in by_sheet.items():
        # Послідовність унікальних груп на цьому листі (у порядку появи)
        seen = []
        for r in rows:
            g = (r.get("group") or "").strip()
            if g and g not in seen:
                seen.append(g)
        for idx, g in enumerate(seen):
            positions[g].append(idx)

    # Сортуємо за median position; tie-breaker — за середньою позицією
    def rank_key(g):
        ps = positions[g]
        return (median(ps), sum(ps) / len(ps), g)

    return sorted(positions.keys(), key=rank_key)


def aggregate_articles_in_group(by_sheet: dict, group: str) -> list[str]:
    """Для конкретної групи збирає позиції статей у межах цієї групи на кожному листі.
    Повертає список статей у порядку median-позиції.
    """
    positions = defaultdict(list)  # article_name → [pos_in_group]

    for sheet_name, rows in by_sheet.items():
        # Статті цієї групи в порядку появи
        seen = []
        for r in rows:
            if (r.get("group") or "").strip() != group:
                continue
            a = (r.get("article") or "").strip()
            if a and a not in seen:
                seen.append(a)
        for idx, a in enumerate(seen):
            positions[a].append(idx)

    def rank_key(a):
        ps = positions[a]
        return (median(ps), sum(ps) / len(ps), a)

    return sorted(positions.keys(), key=rank_key)


def main():
    # 1. Знаходимо Excel-файл за міткою "Лютий 2026"
    cfg_excel = next(e for e in config.EXCEL_FILES if e["label"] == EXCEL_LABEL)
    xlsx_path = cfg_excel["path"]
    month_header = cfg_excel["month_header"]

    # 2. Список листів з 1С
    print(f"Connecting to ERP for sheet list (doc №{DOC_NUMBER})...")
    conn = connect_erp()
    used_sheets = get_used_sheets(conn)
    print(f"  Found {len(used_sheets)} sheets in А_РасшифровкаЛистов")

    # 3. Парсимо Excel
    print(f"Parsing Excel: {xlsx_path}")
    by_sheet = load_excel_sheets(xlsx_path, month_header, set(used_sheets))
    matched = sorted(by_sheet.keys())
    not_found_in_excel = [s for s in used_sheets if s not in by_sheet]
    print(f"  Matched in Excel: {len(matched)} / {len(used_sheets)}")
    if not_found_in_excel:
        print(f"  Sheets not found in Excel ({len(not_found_in_excel)}):")
        for s in not_found_in_excel[:10]:
            print(f"    - {s}")
        if len(not_found_in_excel) > 10:
            print(f"    ... and {len(not_found_in_excel) - 10} more")

    # 4. Агрегуємо порядок груп
    group_order = aggregate_group_order(by_sheet)
    groups_json = [{"name": g, "sort": (i + 1) * 100} for i, g in enumerate(group_order)]
    print(f"\nGroups ({len(groups_json)}):")
    for g in groups_json:
        print(f"  {g['sort']:>5} | {g['name']}")

    # 5. Агрегуємо порядок статей в межах кожної групи
    articles_json = []
    print("\nArticles by group:")
    for g in group_order:
        articles = aggregate_articles_in_group(by_sheet, g)
        print(f"  [{g}] — {len(articles)} articles")
        for i, a in enumerate(articles):
            articles_json.append({
                "name": a,
                "group": g,
                "sort": (i + 1) * 100,
            })

    # 6. JSON
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "source_file": xlsx_path,
        "source_doc": f"А_РасшифровкаЛистов №{DOC_NUMBER} від {DOC_DATE.date().isoformat()}",
        "sheets_used": used_sheets,
        "sheets_matched_in_excel": matched,
        "sheets_not_found_in_excel": not_found_in_excel,
        "generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "numbering_scheme": "step=100, restart per group for articles",
        "groups": groups_json,
        "articles": articles_json,
    }
    OUT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nSaved: {OUT_JSON}")
    print(f"  groups={len(groups_json)} articles={len(articles_json)}")


if __name__ == "__main__":
    main()
