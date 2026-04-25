"""Синхронизация JSON-мапингов с Справочник.А_Статьи_PL (1С — источник истины для защищённых статей).

Для каждой статьи, помеченной в 1С как «Ручная корректировка» (А_РучнаяКорректировка = Истина),
скрипт подтягивает актуальное состояние из справочника и перезаписывает JSON-файлы пайплайна:

  05_mapping_pl_to_dds.json        dds_items (ТЧ Статьи)       match_type → catalog_lock
  10_mapping_statya_dohodov.json   statya_dohodov_uuid/name    match_type → catalog_lock
  02_unique_articles.json          uuid, group, type_statii

Зачем: 11_upload_articles.py уже защищает 1С от перезаписи (флаг пропускает запись), но JSON-файлы
генерируются из Excel+fuzzy и не отражают ручные правки. После запуска 15_sync_*.py пайплайн будет
видеть тот же мапинг, что финансист сконфигурировал в 1С, а отчёты / документы А_ОтчетPL
будут строиться на актуальных связях.

Рекомендуемый момент запуска:
  (1) После серии ручных правок в 1С и перед следующим прогоном 08_prepare_documents.py.
  (2) Можно включить в канонический порядок после 04_match_articles.py / 04b_match_statya_dohodov.py.

Usage:
  python scripts/15_sync_catalog_to_json.py --dry-run
  python scripts/15_sync_catalog_to_json.py
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str


QUERY_LOCKED = """
ВЫБРАТЬ
    PL.Ссылка                        КАК Ссылка,
    PL.Наименование                  КАК Наименование,
    PL.Код                           КАК Код,
    PL.Группа                        КАК Группа,
    PL.Группа.Наименование           КАК ГруппаНаименование,
    PL.СтатьяДвиженияДенежныхСредств КАК ДДСШапка,
    PL.СтатьяДоходов                 КАК СтатьяДоходов,
    PL.ТипСтатьи                     КАК ТипСтатьи
ИЗ Справочник.А_Статьи_PL КАК PL
ГДЕ НЕ PL.ЭтоГруппа
    И НЕ PL.ПометкаУдаления
    И PL.А_РучнаяКорректировка
"""

QUERY_TCH = """
ВЫБРАТЬ
    ТЧ.Ссылка                          КАК PLСсылка,
    ТЧ.НомерСтроки                     КАК НомерСтроки,
    ТЧ.СтатьяДвиженияДенежныхСредств   КАК ДДС,
    ТЧ.СтатьяДвиженияДенежныхСредств.Наименование КАК ДДСИмя
ИЗ Справочник.А_Статьи_PL.Статьи КАК ТЧ
ГДЕ ТЧ.Ссылка В (&Ссылки)
УПОРЯДОЧИТЬ ПО ТЧ.Ссылка, ТЧ.НомерСтроки
"""


def _fetch_locked_state(conn):
    """Возвращает dict {pl_name: {...}} со всеми защищёнными статьями."""
    q = conn.NewObject("Запрос")
    q.Текст = QUERY_LOCKED
    tz = q.Выполнить().Выгрузить()

    state = {}
    refs = []
    for i in range(tz.Количество()):
        r = tz.Получить(i)
        ref = r.Ссылка
        refs.append(ref)

        name = str(r.Наименование)
        group_name = str(r.ГруппаНаименование) if r.ГруппаНаименование else ""
        # ТипСтатьи → строковое имя значения перечисления
        type_statii = ""
        try:
            type_statii = str(conn.XMLСтрока(r.ТипСтатьи)) if r.ТипСтатьи else ""
        except Exception:
            pass
        # Форматы XMLСтрока для перечислений не всегда нужны — используем XMLString виртуал только
        # чтобы извлечь имя значения; проще — через КакойEnumValue.
        # Нормализуем: «Доход», «Расход», «ОперационныйИтог», «Информационный»
        try:
            # Сравниваем с перечислением напрямую (robust метод)
            if r.ТипСтатьи == conn.Перечисления.А_ТипСтатьиPL.Доход:
                type_statii = "Доход"
            elif r.ТипСтатьи == conn.Перечисления.А_ТипСтатьиPL.Расход:
                type_statii = "Расход"
            elif r.ТипСтатьи == conn.Перечисления.А_ТипСтатьиPL.ОперационныйИтог:
                type_statii = "ОперационныйИтог"
            elif r.ТипСтатьи == conn.Перечисления.А_ТипСтатьиPL.Информационный:
                type_statii = "Информационный"
        except Exception:
            pass

        sd_uuid = ""
        sd_name = ""
        if r.СтатьяДоходов and not r.СтатьяДоходов.Пустая():
            sd_uuid = uuid_str(conn, r.СтатьяДоходов)
            sd_name = str(r.СтатьяДоходов.Наименование)

        state[name] = {
            "ref": ref,
            "uuid": uuid_str(conn, ref),
            "code": str(r.Код).strip(),
            "group_name": group_name,
            "group_uuid": uuid_str(conn, r.Группа) if r.Группа and not r.Группа.Пустая() else "",
            "type_statii": type_statii,
            "statya_dohodov_uuid": sd_uuid,
            "statya_dohodov_name": sd_name,
            "dds_items": [],  # заполним ниже
        }

    # Вторая выборка — ТЧ Статьи (DDS list) по всем ссылкам одним запросом
    if refs:
        q2 = conn.NewObject("Запрос")
        q2.Текст = QUERY_TCH
        arr = conn.NewObject("Массив")
        for ref in refs:
            arr.Добавить(ref)
        q2.УстановитьПараметр("Ссылки", arr)
        tz2 = q2.Выполнить().Выгрузить()

        by_ref_uuid = {uuid_str(conn, ref): ref for ref in refs}
        name_by_ref_uuid = {
            uuid_str(conn, v["ref"]): k for k, v in state.items()
        }

        for i in range(tz2.Количество()):
            r = tz2.Получить(i)
            pl_ref_uuid = uuid_str(conn, r.PLСсылка)
            pl_name = name_by_ref_uuid.get(pl_ref_uuid)
            if not pl_name:
                continue
            if not r.ДДС or r.ДДС.Пустая():
                continue
            state[pl_name]["dds_items"].append({
                "uuid": uuid_str(conn, r.ДДС),
                "name": str(r.ДДСИмя),
            })

    return state


def _update_05_mapping(locked_state, dry_run):
    """Обновить 05_mapping_pl_to_dds.json: dds_items ← ТЧ Статьи из 1С для locked статей."""
    path = config.JSON_DIR / "05_mapping_pl_to_dds.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    mappings = data["mappings"]

    changed = []
    for m in mappings:
        nm = m["pl_name"]
        if nm not in locked_state:
            continue
        new_items = [
            {"uuid": d["uuid"], "name": d["name"], "score": 100, "auto": True, "source": "catalog_lock"}
            for d in locked_state[nm]["dds_items"]
        ]
        old_items = m.get("dds_items", [])
        # сравнение по списку uuid (без порядка)
        if {d["uuid"] for d in new_items} != {d.get("uuid") for d in old_items} \
                or len(new_items) != len(old_items) \
                or m.get("match_type") != "catalog_lock":
            changed.append({
                "pl_name": nm,
                "old_uuids": [d.get("uuid") for d in old_items],
                "new_uuids": [d["uuid"] for d in new_items],
            })
            m["dds_items"] = new_items
            m["candidates"] = []
            m["match_type"] = "catalog_lock"

    if not dry_run:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return changed, path


def _update_10_mapping(locked_state, dry_run):
    """Обновить 10_mapping_statya_dohodov.json: statya_dohodov_uuid ← из 1С для locked статей."""
    path = config.JSON_DIR / "10_mapping_statya_dohodov.json"
    if not path.exists():
        return [], path
    data = json.loads(path.read_text(encoding="utf-8"))
    mappings = data.get("mappings", [])
    existing_by_name = {m["pl_name"]: m for m in mappings}

    changed = []
    # Обновить существующие + добавить новые при наличии statya_dohodov в 1С
    for name, st in locked_state.items():
        sd_uuid = st["statya_dohodov_uuid"]
        if not sd_uuid:
            # нет СтатьиДоходов в 1С → если запись есть в JSON, очистим
            if name in existing_by_name:
                m = existing_by_name[name]
                if m.get("statya_dohodov_uuid"):
                    changed.append({"pl_name": name, "old": m.get("statya_dohodov_uuid"), "new": None})
                    m["statya_dohodov_uuid"] = ""
                    m["statya_dohodov_name"] = ""
                    m["match_type"] = "catalog_lock_empty"
                    m["candidates"] = []
            continue

        if name in existing_by_name:
            m = existing_by_name[name]
            if m.get("statya_dohodov_uuid") != sd_uuid or m.get("match_type") != "catalog_lock":
                changed.append({"pl_name": name, "old": m.get("statya_dohodov_uuid"), "new": sd_uuid})
                m["statya_dohodov_uuid"] = sd_uuid
                m["statya_dohodov_name"] = st["statya_dohodov_name"]
                m["match_type"] = "catalog_lock"
                m["candidates"] = []
        else:
            new_m = {
                "pl_name": name,
                "statya_dohodov_uuid": sd_uuid,
                "statya_dohodov_name": st["statya_dohodov_name"],
                "match_type": "catalog_lock",
                "candidates": [],
            }
            mappings.append(new_m)
            changed.append({"pl_name": name, "old": None, "new": sd_uuid})

    data["total"] = len(mappings)
    data["mappings"] = mappings
    if not dry_run:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return changed, path


def _update_02_articles(locked_state, dry_run):
    """Обновить 02_unique_articles.json: uuid, group, type_statii ← из 1С для locked статей."""
    path = config.JSON_DIR / "02_unique_articles.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    articles = data["articles"]

    changed = []
    for a in articles:
        nm = a.get("name")
        if nm not in locked_state:
            continue
        st = locked_state[nm]
        diff = {}
        if a.get("uuid") != st["uuid"]:
            diff["uuid"] = (a.get("uuid"), st["uuid"])
            a["uuid"] = st["uuid"]
        # group — в JSON строковое имя
        if st["group_name"] and a.get("group") != st["group_name"]:
            diff["group"] = (a.get("group"), st["group_name"])
            a["group"] = st["group_name"]
        # type_statii
        if st["type_statii"] and a.get("type_statii") != st["type_statii"]:
            diff["type_statii"] = (a.get("type_statii"), st["type_statii"])
            a["type_statii"] = st["type_statii"]
        if diff:
            changed.append({"pl_name": nm, **diff})

    if not dry_run:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return changed, path


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dry-run", action="store_true", help="только показать diff, не писать JSON")
    args = parser.parse_args()

    print("Подключаюсь к ЕРП и читаю защищённые (А_РучнаяКорректировка=Истина) статьи…")
    conn = connect_erp()
    locked_state = _fetch_locked_state(conn)
    print(f"Защищённых статей в 1С: {len(locked_state)}")
    empty_tch = [n for n, s in locked_state.items() if not s["dds_items"]]
    if empty_tch:
        print(f"  (!) {len(empty_tch)} статей БЕЗ ДДС в ТЧ: {', '.join(empty_tch[:5])}{'…' if len(empty_tch)>5 else ''}")

    print("\n=== 05_mapping_pl_to_dds.json ===")
    changes_05, path_05 = _update_05_mapping(locked_state, args.dry_run)
    print(f"  Изменений: {len(changes_05)}")
    for c in changes_05[:15]:
        print(f"    • {c['pl_name']}: {c['old_uuids']} → {c['new_uuids']}")
    if len(changes_05) > 15:
        print(f"    … и ещё {len(changes_05)-15}")

    print("\n=== 10_mapping_statya_dohodov.json ===")
    changes_10, path_10 = _update_10_mapping(locked_state, args.dry_run)
    print(f"  Изменений: {len(changes_10)}")
    for c in changes_10:
        print(f"    • {c['pl_name']}: {c['old']} → {c['new']}")

    print("\n=== 02_unique_articles.json ===")
    changes_02, path_02 = _update_02_articles(locked_state, args.dry_run)
    print(f"  Изменений: {len(changes_02)}")
    for c in changes_02[:15]:
        diff_fmt = "; ".join(f"{k}: {v[0]!r} → {v[1]!r}" for k, v in c.items() if k != "pl_name")
        print(f"    • {c['pl_name']}: {diff_fmt}")
    if len(changes_02) > 15:
        print(f"    … и ещё {len(changes_02)-15}")

    if args.dry_run:
        print("\n[DRY-RUN] Ничего не записано.")
    else:
        print(f"\nЗаписаны:")
        print(f"  {path_05}")
        print(f"  {path_10}")
        print(f"  {path_02}")


if __name__ == "__main__":
    main()
