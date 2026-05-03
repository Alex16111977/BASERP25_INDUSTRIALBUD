"""Step 08: подготовить 08_documents_to_import.json на основе
- 01_raw_sheets.json (данные листов Excel — все visible PnL-листы)
- Документ.А_РасшифровкаЛистов (маппинг лист -> подразделение/направление/ВключатьДочерние; источник истины UI 1С).

Алгоритм:
1. Для каждого Excel-периода (Декабрь/Январь/Лютий) найти соответствующий документ А_РасшифровкаЛистов по дате (конец месяца 12:00:00).
2. Из ТЧ Расшифровка построить mapping sheet_name -> {podrazdelenie, direction, include_children}.
3. Пройтись по листам Excel:
   - Только листы у которых в mapping заполнено Подразделение → формируют документ А_ОтчетPL.
   - Остальные (без подразделения в A_РасшифровкаЛистов) пропускаются.
4. Записать 08_documents_to_import.json.
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str
from utils.names_cleaner import normalize


def load_rasshifrovka_mapping(conn):
    """Для каждого документа А_РасшифровкаЛистов возвращает {date_iso: {sheet_name: {...}}}."""
    q = conn.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ Ссылка
    ИЗ Документ.А_РасшифровкаЛистов
    ГДЕ НЕ ПометкаУдаления
    УПОРЯДОЧИТЬ ПО Дата
    """
    tz = q.Выполнить().Выгрузить()
    by_date = {}
    for i in range(tz.Количество()):
        ref = tz.Получить(i).Ссылка
        obj = ref.ПолучитьОбъект()
        date_iso = str(obj.Дата)[:10]  # YYYY-MM-DD
        sheets = {}
        for row in obj.Расшифровка:
            podr = row.Подразделение
            podr_empty = podr.Пустая() if podr else True
            napr = row.НаправлениеДеятельности
            napr_empty = napr.Пустая() if napr else True
            sn = str(row.ИмяЛиста)
            sheets[sn] = {
                "struct_uuid": None if podr_empty else uuid_str(conn, podr),
                "struct_name": None if podr_empty else str(podr.Наименование),
                "direction_uuid": None if napr_empty else uuid_str(conn, napr),
                "direction_name": None if napr_empty else str(napr.Наименование),
                "include_children": bool(row.ВключатьДочерние),
                "podr_stroka": str(row.ПодразделениеСтрока),
            }
        by_date[date_iso] = {
            "doc_number": str(obj.Номер).strip(),
            "sheets": sheets,
        }
    return by_date


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", default=None,
                    help="Опционально: YYYY-MM. Включить в 08_documents_to_import только периоды этого месяца.")
    args = ap.parse_args()
    month_filter = args.month
    if month_filter:
        print(f"[FILTER] обрабатываем только месяц {month_filter}")

    raw = json.loads((config.JSON_DIR / "01_raw_sheets.json").read_text(encoding="utf-8"))
    arts = json.loads((config.JSON_DIR / "02_unique_articles.json").read_text(encoding="utf-8"))["articles"]
    groups = json.loads((config.JSON_DIR / "03_groups_pl.json").read_text(encoding="utf-8"))["groups"]
    # Fallback для направления: если в ТЧ Расшифровка направление не указано — брать из подразделения (06)
    struct = json.loads((config.JSON_DIR / "06_struct_predpr.json").read_text(encoding="utf-8"))["items"]
    struct_by_uuid = {s["uuid"]: s for s in struct}
    group_uuid_by_name = {g["name"]: g.get("uuid", "") for g in groups}
    group_uuid_by_norm = {normalize(g["name"]): g.get("uuid", "") for g in groups}

    # Lookup статей по нормализованному имени (с учётом алиасов опечаток из config.py).
    aliases = getattr(config, "ARTICLE_NAME_ALIASES", {}) or {}
    by_norm = {}
    for a in arts:
        by_norm[normalize(a["name"])] = a
        for v in a.get("variants", []):
            by_norm[normalize(v)] = a
    # Добавляем нормализованные алиасы как ключи на каноничную статью.
    for src_name, canon_name in aliases.items():
        canon_a = next((a for a in arts if a["name"] == canon_name), None)
        if canon_a:
            by_norm[normalize(src_name)] = canon_a

    # Маппинг из А_РасшифровкаЛистов (источник истины)
    print("Читаем Документ.А_РасшифровкаЛистов (маппинг из 1С)...")
    conn = connect_erp()
    rasshifrovka = load_rasshifrovka_mapping(conn)
    for date_iso, info in rasshifrovka.items():
        print(f"  {info['doc_number']} ({date_iso}): {len(info['sheets'])} листов в ТЧ")

    docs = []
    missing_articles = []
    skipped_no_mapping = []
    skipped_no_podr = []
    total_rows = 0

    for period in raw:
        date_iso = period["period"]
        if month_filter and not date_iso.startswith(month_filter):
            continue
        # ИСКАТЬ ПО ДАТЕ — Excel period ("2025-12-31") должен совпадать с датой А_Расшифровки.
        if date_iso not in rasshifrovka:
            print(f"  WARN: период {date_iso} не имеет А_РасшифровкаЛистов")
            continue
        mapping_for_period = rasshifrovka[date_iso]["sheets"]

        for sh in period["sheets"]:
            sname = sh["sheet_name"]
            m = mapping_for_period.get(sname)
            if m is None:
                skipped_no_mapping.append((date_iso, sname))
                continue
            if not m.get("struct_uuid"):
                skipped_no_podr.append((date_iso, sname))
                continue

            # Fallback: если финансист в А_Расшифровке не проставил направление, но подразделение
            # имеет А_НаправлениеДеятельности — подтянуть из 06_struct_predpr.
            if not m.get("direction_uuid"):
                fallback = struct_by_uuid.get(m["struct_uuid"])
                if fallback and fallback.get("direction_uuid"):
                    m["direction_uuid"] = fallback["direction_uuid"]
                    m["direction_name"] = fallback.get("direction_name") or ""

            rows_out = []
            for r in sh["rows"]:
                art = by_norm.get(normalize(r["article"]))
                if not art:
                    missing_articles.append((r["article"], sname))
                    continue
                rows_out.append({
                    "article_name": art["name"],
                    "article_uuid": art.get("uuid", ""),
                    "group_name": r.get("group", "") or art.get("group", ""),
                    "type_statii": r.get("type_statii") or art.get("type_statii") or "Расход",
                    "row_excel": r.get("row", 0),
                    "comment": r.get("comment", ""),
                    "sum_f1": r["sum_f1"],
                    "sum_f2": r["sum_f2"],
                    "total": r["total"],
                })

            # Group totals
            group_totals_out = []
            for gt in sh.get("group_totals", []):
                gname = gt["group_name"]
                guid = group_uuid_by_name.get(gname) or group_uuid_by_norm.get(normalize(gname), "")
                group_totals_out.append({
                    "group_name": gname,
                    "group_uuid": guid,
                    "sum_f1_excel": gt["sum_f1_excel"],
                    "sum_f2_excel": gt["sum_f2_excel"],
                    "sum_excel": gt["sum_excel"],
                })

            general_totals_out = []
            for gt in sh.get("general_totals", []):
                general_totals_out.append({
                    "show": gt["show"],
                    "vid": gt["vid"],
                    "sum_f1": gt["sum_f1"],
                    "sum_f2": gt["sum_f2"],
                    "sum_total": gt["sum_total"],
                })

            docs.append({
                "date": date_iso,
                "period_label": period["label"],
                "organization": config.ORGANIZATION_NAME,
                "sheet_name": sname,
                "podrazdelenie_uuid": m["struct_uuid"] or "",
                "podrazdelenie_name": m["struct_name"] or "",
                "podrazdelenie_stroka": (m.get("podr_stroka") or sname).strip(),
                "direction_uuid": m.get("direction_uuid") or "",
                "direction_name": m.get("direction_name") or "",
                "include_children": bool(m.get("include_children", False)),
                "rows": rows_out,
                "group_totals": group_totals_out,
                "general_totals": general_totals_out,
            })
            total_rows += len(rows_out)

    dst = config.JSON_DIR / "08_documents_to_import.json"
    dst.write_text(
        json.dumps({
            "total_docs": len(docs),
            "total_rows": total_rows,
            "skipped_no_mapping": len(skipped_no_mapping),
            "skipped_no_podr": len(skipped_no_podr),
            "skipped_no_mapping_list": skipped_no_mapping,
            "skipped_no_podr_list": skipped_no_podr,
            "documents": docs,
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"\nWrote {dst}  (docs: {len(docs)}, rows: {total_rows})")
    print(f"  skipped (нет в А_РасшифровкаЛистов): {len(skipped_no_mapping)}")
    print(f"  skipped (нет Подразделения в А_Расшифровке): {len(skipped_no_podr)}")
    if missing_articles:
        print(f"  WARN: {len(missing_articles)} rows with missing articles in catalog")
        for nm, sh in missing_articles[:5]:
            print(f"    - {sh} :: {nm}")


if __name__ == "__main__":
    main()
