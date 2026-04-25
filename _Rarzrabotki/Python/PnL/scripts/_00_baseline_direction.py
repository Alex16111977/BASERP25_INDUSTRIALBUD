"""Phase 0 baseline: snapshot текущего состояния перед миграцией направления.

Собирает:
- 3 документа А_РасшифровкаЛистов (Номер, Дата, кол-во строк, с подразделением)
- Счётчик А_ОтчетPL (всего, с подразделением)
- 8 направлений (UUID + имя)
- Покрытие направления по struct_uuid из 07_mapping: сколько из 27/23 имеют А_НаправлениеДеятельности

Пишет результат в data/json/00_baseline_pre_direction_migration.json.
Read-only: базу не меняет.
"""
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str


def main():
    conn = connect_erp()

    # 1) Документы А_РасшифровкаЛистов
    q = conn.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ Д.Номер, Д.Дата,
        КОЛИЧЕСТВО(ТЧ.НомерСтроки) КАК Строк,
        СУММА(ВЫБОР КОГДА ТЧ.Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
            ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СПодразд
    ИЗ Документ.А_РасшифровкаЛистов КАК Д
    ЛЕВОЕ СОЕДИНЕНИЕ Документ.А_РасшифровкаЛистов.Расшифровка КАК ТЧ
        ПО ТЧ.Ссылка = Д.Ссылка
    ГДЕ НЕ Д.ПометкаУдаления
    СГРУППИРОВАТЬ ПО Д.Номер, Д.Дата
    УПОРЯДОЧИТЬ ПО Д.Дата
    """
    tz = q.Выполнить().Выгрузить()
    rasshifrovka = []
    for i in range(tz.Количество()):
        r = tz.Получить(i)
        rasshifrovka.append({
            "number": str(r.Номер).strip(),
            "date": str(r.Дата)[:10],
            "rows": int(r.Строк),
            "with_podrazd": int(r.СПодразд),
        })

    # 2) А_ОтчетPL summary
    q2 = conn.NewObject("Запрос")
    q2.Текст = """
    ВЫБРАТЬ
        КОЛИЧЕСТВО(Д.Ссылка) КАК Всего,
        СУММА(ВЫБОР КОГДА Д.ПометкаУдаления ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК ПометкаУд,
        СУММА(ВЫБОР КОГДА Д.Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
            ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СПодразд
    ИЗ Документ.А_ОтчетPL КАК Д
    """
    tz2 = q2.Выполнить().Выгрузить()
    r2 = tz2.Получить(0)
    otchet_pl = {
        "total": int(r2.Всего),
        "pometka_ud": int(r2.ПометкаУд),
        "with_podrazd": int(r2.СПодразд),
    }

    # 3) Направления деятельности
    q3 = conn.NewObject("Запрос")
    q3.Текст = """
    ВЫБРАТЬ Н.Ссылка, Н.Наименование
    ИЗ Справочник.НаправленияДеятельности КАК Н
    ГДЕ НЕ Н.ПометкаУдаления
    УПОРЯДОЧИТЬ ПО Н.Наименование
    """
    tz3 = q3.Выполнить().Выгрузить()
    napravleniya = []
    for i in range(tz3.Количество()):
        r = tz3.Получить(i)
        napravleniya.append({
            "uuid": uuid_str(conn, r.Ссылка),
            "name": str(r.Наименование),
        })

    # 4) Покрытие направления для struct_uuid из 07_mapping
    mapping_data = json.loads(
        (config.JSON_DIR / "07_mapping_sheet_to_struct.json").read_text(encoding="utf-8")
    )
    mappings = mapping_data["mappings"]
    unique_uuids = sorted({m["struct_uuid"] for m in mappings if m.get("struct_uuid")})

    direction_coverage = []
    with_direction = 0
    without_direction = 0
    for su in unique_uuids:
        try:
            uid = conn.NewObject("УникальныйИдентификатор", su)
            ref = conn.Справочники.СтруктураПредприятия.ПолучитьСсылку(uid)
            name = str(ref.Наименование) if ref else ""
            napr = ref.А_НаправлениеДеятельности if ref else None
            napr_empty = napr.Пустая() if napr else True
            napr_name = str(napr.Наименование) if (napr and not napr_empty) else None
            direction_coverage.append({
                "struct_uuid": su,
                "struct_name": name,
                "direction_name": napr_name,
                "has_direction": not napr_empty,
            })
            if napr_empty:
                without_direction += 1
            else:
                with_direction += 1
        except Exception as ex:
            direction_coverage.append({
                "struct_uuid": su,
                "error": str(ex),
            })
            without_direction += 1

    # 5) Write JSON
    out = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "mode": "baseline_pre_direction_migration",
        "rasshifrovka_listov": rasshifrovka,
        "otchet_pl": otchet_pl,
        "napravleniya": napravleniya,
        "napravleniya_count": len(napravleniya),
        "mapping_07_unique_struct_uuids": len(unique_uuids),
        "direction_coverage_with": with_direction,
        "direction_coverage_without": without_direction,
        "direction_coverage": direction_coverage,
    }
    dst = config.JSON_DIR / "00_baseline_pre_direction_migration.json"
    dst.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"=== Phase 0 baseline ===")
    print(f"А_РасшифровкаЛистов: {len(rasshifrovka)} документов")
    for d in rasshifrovka:
        print(f"  {d['number']} ({d['date']}): {d['rows']} строк, {d['with_podrazd']} с подразд")
    print(f"А_ОтчетPL: всего={otchet_pl['total']}, пометка_уд={otchet_pl['pometka_ud']}, с_подразд={otchet_pl['with_podrazd']}")
    print(f"Направлений деятельности: {len(napravleniya)}")
    for n in napravleniya:
        print(f"  {n['uuid']}  {n['name']}")
    print(f"Mapping 07: {len(unique_uuids)} уникальных struct_uuid")
    print(f"  с направлением: {with_direction}")
    print(f"  без направления: {without_direction}")
    if without_direction > 0:
        print("  Подразделения без направления:")
        for d in direction_coverage:
            if not d.get("has_direction"):
                print(f"    - {d.get('struct_name', '?')} ({d.get('struct_uuid', '?')})")
    print(f"\nWrote {dst}")


if __name__ == "__main__":
    main()
