"""Step 06: COM → 06_struct_predpr.json (dump СтруктураПредприятия)."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str


QUERY = """
ВЫБРАТЬ
    СП.Ссылка КАК Ссылка,
    СП.Наименование КАК Наименование,
    СП.Код КАК Код,
    СП.Родитель КАК Родитель,
    СП.Родитель.Наименование КАК РодительНаименование,
    СП.А_НаправлениеДеятельности КАК НаправлениеДеятельности,
    СП.А_НаправлениеДеятельности.Наименование КАК НаправлениеИмя,
    СП.А_ЭтоПодразделениеНаправление КАК ЭтоНаправление,
    СП.ПометкаУдаления КАК ПометкаУдаления
ИЗ
    Справочник.СтруктураПредприятия КАК СП
ГДЕ
    НЕ СП.ПометкаУдаления
УПОРЯДОЧИТЬ ПО
    СП.Наименование
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
        napr_ref = r.НаправлениеДеятельности
        napr_empty = napr_ref.Пустая() if napr_ref else True
        parent_ref = r.Родитель
        parent_empty = parent_ref.Пустая() if parent_ref else True
        out.append({
            "uuid": uuid_str(conn, r.Ссылка),
            "name": str(r.Наименование),
            "code": str(r.Код),
            "parent_uuid": None if parent_empty else uuid_str(conn, parent_ref),
            "parent_name": str(r.РодительНаименование) if r.РодительНаименование else "",
            "direction_uuid": None if napr_empty else uuid_str(conn, napr_ref),
            "direction_name": str(r.НаправлениеИмя) if r.НаправлениеИмя else "",
            "is_direction": bool(r.ЭтоНаправление),
        })
    dst = config.JSON_DIR / "06_struct_predpr.json"
    dst.write_text(json.dumps({"total": len(out), "items": out}, ensure_ascii=False, indent=2),
                   encoding="utf-8")
    print(f"Wrote {dst}  (items: {len(out)})")


if __name__ == "__main__":
    main()
