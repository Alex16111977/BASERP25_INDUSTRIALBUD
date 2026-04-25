"""Прочитать документ А_РасшифровкаЛистов №000000002 (эталон от финансиста)
и вывести JSON с полным маппингом sheet_name -> struct + direction + include_children."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp, uuid_str


def main():
    conn = connect_erp()
    q = conn.NewObject("Запрос")
    import os
    doc_num = os.environ.get("DOC_NUM", "000000002")
    q.Текст = f"""
    ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
    ИЗ Документ.А_РасшифровкаЛистов
    ГДЕ Номер = "{doc_num}" И НЕ ПометкаУдаления
    """
    tz = q.Выполнить().Выгрузить()
    if tz.Количество() == 0:
        print("ERR: doc not found")
        return
    ref = tz.Получить(0).Ссылка
    obj = ref.ПолучитьОбъект()

    rows = []
    for row in obj.Расшифровка:
        podr_ref = row.Подразделение
        podr_empty = podr_ref.Пустая() if podr_ref else True
        napr_ref = row.НаправлениеДеятельности
        napr_empty = napr_ref.Пустая() if napr_ref else True
        rows.append({
            "n": len(rows) + 1,
            "sheet_name": str(row.ИмяЛиста),
            "podr_stroka": str(row.ПодразделениеСтрока),
            "struct_uuid": None if podr_empty else uuid_str(conn, podr_ref),
            "struct_name": None if podr_empty else str(podr_ref.Наименование),
            "include_children": bool(row.ВключатьДочерние),
            "direction_uuid": None if napr_empty else uuid_str(conn, napr_ref),
            "direction_name": None if napr_empty else str(napr_ref.Наименование),
        })

    out_path = Path(__file__).resolve().parents[1] / "data" / "json" / f"_doc{doc_num}_reference_mapping.json"
    out_path.write_text(
        json.dumps({"doc": doc_num, "rows": rows}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Wrote {out_path}  ({len(rows)} rows)")
    # Print summary
    incch = sum(1 for r in rows if r["include_children"])
    with_dir = sum(1 for r in rows if r["direction_uuid"])
    with_podr = sum(1 for r in rows if r["struct_uuid"])
    print(f"total={len(rows)} with_podr={with_podr} with_dir={with_dir} include_children={incch}")


if __name__ == "__main__":
    main()
