"""Step 02b: подтянуть UUID существующих элементов Справочник.А_Статьи_PL
в 02_unique_articles.json (без создания новых элементов).
Нужен после 02 перед запуском 08 (08 требует uuid для ТЧ ДанныеОтчета)."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str


def main():
    src = config.JSON_DIR / "02_unique_articles.json"
    data = json.loads(src.read_text(encoding="utf-8"))
    articles = data["articles"]

    conn = connect_erp()
    mgr = conn.Справочники.А_Статьи_PL

    found = not_found = 0
    not_found_list = []
    for a in articles:
        # 1) По каноническому имени
        ref = mgr.НайтиПоНаименованию(a["name"], True)
        if ref is None or ref.Пустая():
            # 2) По вариантам (если есть другие написания)
            for v in a.get("variants", []):
                ref = mgr.НайтиПоНаименованию(v, True)
                if ref and not ref.Пустая():
                    break
        if ref and not ref.Пустая():
            a["uuid"] = uuid_str(conn, ref)
            found += 1
        else:
            a["uuid"] = ""
            not_found += 1
            not_found_list.append(a["name"])

    src.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Total: {len(articles)}, found UUID: {found}, not_found: {not_found}")
    if not_found_list:
        print("Not found in catalog А_Статьи_PL (нужно создать вручную или через 11):")
        for nm in not_found_list[:20]:
            print(f"  - {nm}")


if __name__ == "__main__":
    main()
