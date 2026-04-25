"""Step 03: COM → 04_dds_articles.json (dump СтатьиДвиженияДенежныхСредств)."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str


QUERY = """
ВЫБРАТЬ
    СДДС.Ссылка КАК Ссылка,
    СДДС.Наименование КАК Наименование,
    СДДС.Код КАК Код,
    СДДС.Родитель.Наименование КАК РодительНаименование,
    СДДС.ЭтоГруппа КАК ЭтоГруппа,
    СДДС.ПометкаУдаления КАК ПометкаУдаления,
    СДДС.А_НеИспользоватьКазна КАК НеИспользовать
ИЗ
    Справочник.СтатьиДвиженияДенежныхСредств КАК СДДС
ГДЕ
    НЕ СДДС.ПометкаУдаления
УПОРЯДОЧИТЬ ПО
    СДДС.Наименование
"""


def main():
    conn = connect_erp()
    q = conn.NewObject("Запрос")
    q.Текст = QUERY
    tz = q.Выполнить().Выгрузить()
    out = []
    n = tz.Количество()
    for i in range(n):
        r = tz.Получить(i)
        out.append({
            "uuid": uuid_str(conn, r.Ссылка),
            "name": str(r.Наименование),
            "code": str(r.Код),
            "parent": str(r.РодительНаименование) if r.РодительНаименование else "",
            "is_group": bool(r.ЭтоГруппа),
        })
    dst = config.JSON_DIR / "04_dds_articles.json"
    dst.write_text(json.dumps({"total": len(out), "items": out}, ensure_ascii=False, indent=2),
                   encoding="utf-8")
    print(f"Wrote {dst}  (items: {len(out)})")


if __name__ == "__main__":
    main()
