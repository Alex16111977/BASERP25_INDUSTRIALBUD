"""Step 10: create/find elements in Справочник.А_ГруппаСтатей_PL, write UUIDs back to 03_groups_pl.json."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str


def main():
    src = config.JSON_DIR / "03_groups_pl.json"
    data = json.loads(src.read_text(encoding="utf-8"))
    groups = data["groups"]

    conn = connect_erp()
    mgr = conn.Справочники.А_ГруппаСтатей_PL

    created = found = 0
    for g in groups:
        ref = mgr.НайтиПоНаименованию(g["name"], True)
        if ref and not ref.Пустая():
            g["uuid"] = uuid_str(conn, ref)
            found += 1
            marker = "   "
        else:
            obj = mgr.СоздатьЭлемент()
            obj.Наименование = g["name"]
            obj.Записать()
            g["uuid"] = uuid_str(conn, obj.Ссылка)
            created += 1
            marker = "NEW"
        print(f"  {marker} {g['name']}  ({g['uuid']})")

    src.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Done. found={found} created={created}. Updated {src}")


if __name__ == "__main__":
    main()
