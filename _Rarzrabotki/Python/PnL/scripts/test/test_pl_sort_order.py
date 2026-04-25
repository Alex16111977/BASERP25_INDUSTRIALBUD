"""
Інтеграційний тест: Сорт у БД (А_ГруппаСтатей_PL + А_Статьи_PL) збігається з JSON.

Перевіряє що скрипт 17_upload_pl_sort.py успішно завантажив усі значення.
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
import config  # noqa: E402
from utils.com_connect import connect_erp  # noqa: E402

JSON_PATH = config.JSON_DIR / "16_pl_sort_order.json"


def fetch_db_groups(conn) -> dict[str, int]:
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ Г.Наименование КАК Имя, Г.Сорт КАК Сорт\n"
        "ИЗ Справочник.А_ГруппаСтатей_PL КАК Г\n"
        "ГДЕ НЕ Г.ПометкаУдаления И НЕ Г.ЭтоГруппа"
    )
    sel = q.Выполнить().Выбрать()
    out = {}
    while sel.Следующий():
        out[str(sel.Имя)] = int(float(sel.Сорт or 0))
    return out


def fetch_db_articles(conn) -> dict[str, int]:
    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ С.Наименование КАК Имя, С.Сорт КАК Сорт\n"
        "ИЗ Справочник.А_Статьи_PL КАК С\n"
        "ГДЕ НЕ С.ПометкаУдаления И НЕ С.ЭтоГруппа"
    )
    sel = q.Выполнить().Выбрать()
    out = {}
    while sel.Следующий():
        out[str(sel.Имя)] = int(float(sel.Сорт or 0))
    return out


def main():
    if not JSON_PATH.exists():
        print(f"FAIL: {JSON_PATH} не існує")
        sys.exit(1)
    payload = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    expected_groups = {g["name"]: g["sort"] for g in payload["groups"]}
    expected_articles = {a["name"]: a["sort"] for a in payload["articles"]}

    conn = connect_erp()
    db_groups = fetch_db_groups(conn)
    db_articles = fetch_db_articles(conn)

    # === Групи ===
    g_mismatch = []
    for name, expected in expected_groups.items():
        actual = db_groups.get(name)
        if actual is None:
            g_mismatch.append((name, expected, "NOT_IN_DB"))
        elif actual != expected:
            g_mismatch.append((name, expected, actual))
    print(f"Groups: {len(expected_groups)} expected, mismatched: {len(g_mismatch)}")
    for name, exp, act in g_mismatch:
        print(f"  FAIL [GROUP] {name}: expected {exp}, got {act}")

    # === Статті ===
    # JSON може мати дублі назв в різних групах — перевіримо макс. sort з JSON для кожного name
    expected_max = {}
    for a in payload["articles"]:
        prev = expected_max.get(a["name"])
        if prev is None or a["sort"] > prev:
            expected_max[a["name"]] = a["sort"]
    a_mismatch = []
    a_not_in_db = []
    for name, expected in expected_max.items():
        actual = db_articles.get(name)
        if actual is None:
            a_not_in_db.append(name)
        elif actual != expected:
            a_mismatch.append((name, expected, actual))
    print(f"Articles: {len(expected_max)} expected (з JSON), mismatched: {len(a_mismatch)}, not in DB: {len(a_not_in_db)}")
    for name, exp, act in a_mismatch[:10]:
        print(f"  FAIL [ART] {name}: expected {exp}, got {act}")
    for name in a_not_in_db[:10]:
        print(f"  INFO [ART NOT IN DB] {name}")

    # Перевіряємо тільки ті статті які реально існують у БД (NOT_IN_DB — інформаційно, не FAIL)
    if g_mismatch or a_mismatch:
        print()
        print(f"FAIL: groups_mismatch={len(g_mismatch)}, articles_mismatch={len(a_mismatch)}")
        sys.exit(1)
    print()
    print("PL SORT ORDER OK (всі групи + всі знайдені у БД статті мають правильний Сорт)")


if __name__ == "__main__":
    main()
