"""Step 11: create/find elements in Справочник.А_Статьи_PL, fill ТЧ Статьи with ДДС references."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, is_manually_locked, uuid_str as ref_uuid


def ref_by_uuid(conn, col_name, uuid_value):
    mgr = getattr(conn.Справочники, col_name)
    uid = conn.NewObject("УникальныйИдентификатор", uuid_value)
    ref = mgr.ПолучитьСсылку(uid)
    return ref


def main():
    src_arts = config.JSON_DIR / "02_unique_articles.json"
    src_grp = config.JSON_DIR / "03_groups_pl.json"
    src_map = config.JSON_DIR / "05_mapping_pl_to_dds.json"

    arts_data = json.loads(src_arts.read_text(encoding="utf-8"))
    articles = arts_data["articles"]
    groups = json.loads(src_grp.read_text(encoding="utf-8"))["groups"]
    mapping = json.loads(src_map.read_text(encoding="utf-8"))["mappings"]

    group_uuid_by_name = {g["name"]: g["uuid"] for g in groups if g.get("uuid")}
    dds_by_pl = {m["pl_name"]: m["dds_items"] for m in mapping}

    # СтатьяДоходов маппинг (опционально)
    src_dohod = config.JSON_DIR / "10_mapping_statya_dohodov.json"
    dohod_by_pl = {}
    if src_dohod.exists():
        for m in json.loads(src_dohod.read_text(encoding="utf-8"))["mappings"]:
            if m.get("statya_dohodov_uuid"):
                dohod_by_pl[m["pl_name"]] = m["statya_dohodov_uuid"]

    conn = connect_erp()
    mgr = conn.Справочники.А_Статьи_PL

    created = found = locked = 0
    for a in articles:
        ref = mgr.НайтиПоНаименованию(a["name"], True)
        if ref and not ref.Пустая():
            obj = ref.ПолучитьОбъект()
            # === GUARD: ручная корректировка = наивысший приоритет ===
            # Финансист отметил эту статью как «Ручная корректировка» —
            # пайплайн не трогает ни одно поле, чтобы не затереть ручные правки.
            if is_manually_locked(obj):
                a["uuid"] = ref_uuid(conn, obj.Ссылка)
                locked += 1
                print(f"  LOCK {a['name']}  (А_РучнаяКорректировка=Истина — пропущен)")
                continue
            is_new = False
            found += 1
        else:
            obj = mgr.СоздатьЭлемент()
            obj.Наименование = a["name"]
            is_new = True
            created += 1

        # group
        grp_name = a.get("group") or ""
        grp_uuid = group_uuid_by_name.get(grp_name)
        if grp_uuid:
            obj.Группа = ref_by_uuid(conn, "А_ГруппаСтатей_PL", grp_uuid)

        # ТЧ Статьи (1:N with ДДС)
        obj.Статьи.Очистить()
        first_dds_ref = None
        for dds in dds_by_pl.get(a["name"], []):
            dds_ref = ref_by_uuid(conn, "СтатьиДвиженияДенежныхСредств", dds["uuid"])
            row = obj.Статьи.Добавить()
            row.СтатьяДвиженияДенежныхСредств = dds_ref
            if first_dds_ref is None:
                first_dds_ref = dds_ref

        # Header attribute = first row from ТЧ Статьи
        if first_dds_ref is not None:
            obj.СтатьяДвиженияДенежныхСредств = first_dds_ref

        # Статья доходов (для доходных PL-статей)
        dohod_uuid = dohod_by_pl.get(a["name"])
        if dohod_uuid:
            uid = conn.NewObject("УникальныйИдентификатор", dohod_uuid)
            obj.СтатьяДоходов = conn.ПланыВидовХарактеристик.СтатьиДоходов.ПолучитьСсылку(uid)

        # ТипСтатьи (перечисление А_ТипСтатьиPL)
        t = a.get("type_statii")
        if t:
            try:
                obj.ТипСтатьи = getattr(conn.Перечисления.А_ТипСтатьиPL, t)
            except Exception as ex:
                print(f"  WARN: не удалось установить ТипСтатьи={t!r} у {a['name']!r}: {ex}")

        obj.Записать()
        a["uuid"] = ref_uuid(conn, obj.Ссылка)
        marker = "NEW" if is_new else "   "
        suffix = ""
        if is_new:
            suffix = "   << проверьте и поставьте галку «Ручная корректировка» в 1С!"
        print(f"  {marker} {a['name']}  group={grp_name!r}  dds={len(dds_by_pl.get(a['name'], []))}{suffix}")

    src_arts.write_text(json.dumps(arts_data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Done. found={found} created={created} locked={locked}. Updated {src_arts}")


if __name__ == "__main__":
    main()
