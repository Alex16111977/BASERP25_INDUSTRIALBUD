"""Step 13: fill ТЧ Расшифровка у Документ.А_РасшифровкаЛистов всеми видимыми листами Excel.

- Ищет 3 документа по basename (в базе ИмяФайла хранится только имя файла).
- Для каждого делает backup текущей ТЧ в data/json/13_backup_*.json.
- Очищает ТЧ, заполняет всеми visible-листами.
- Подставляет Подразделение из 07_mapping_sheet_to_struct.json (exact/fuzzy/manual_override).
- Обновляет ИмяФайла до полного пути.
- --dry-run: читает и логирует, но не пишет в базу (backup всё равно делает).
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str

AUTO_MATCH_TYPES = {"exact", "fuzzy", "manual_override"}


def read_visible_and_hidden(excel_path):
    wb = load_workbook(excel_path, read_only=True)
    vis, hid = [], []
    for ws in wb.worksheets:
        (vis if ws.sheet_state == "visible" else hid).append(ws.title)
    wb.close()
    return vis, hid


def find_doc_by_file(conn, excel_path):
    """ИмяФайла — Строка неограниченной длины, сравнивать через ПОДОБНО по basename."""
    basename = Path(excel_path).name
    q = conn.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
    ИЗ Документ.А_РасшифровкаЛистов
    ГДЕ ИмяФайла ПОДОБНО &BasenameLike И НЕ ПометкаУдаления
    УПОРЯДОЧИТЬ ПО Дата УБЫВ"""
    # ПОДОБНО — префиксное: совпадение basename в конце полного пути ИЛИ точно basename
    q.УстановитьПараметр("BasenameLike", f"%{basename}")
    tz = q.Выполнить().Выгрузить()
    if tz.Количество() > 0:
        return tz.Получить(0).Ссылка
    return _find_doc_by_period(conn, excel_path)


def create_doc_for_period(conn, excel_path):
    """Создать новый А_РасшифровкаЛистов на конец месяца excel_file (12:00:00).

    Используется в --month режиме когда документа для месяца нет.
    Возвращает Ссылку на свежесозданный документ (с пустой ТЧ Расшифровка).
    """
    ef = next((e for e in config.EXCEL_FILES if e["path"] == excel_path), None)
    if not ef:
        raise RuntimeError(f"excel_path '{excel_path}' не найден в config.EXCEL_FILES")
    y, m, d = map(int, ef["period"].split("-"))
    dt = datetime(y, m, d, 12, 0, 0)
    obj = conn.Документы.А_РасшифровкаЛистов.СоздатьДокумент()
    obj.Дата = dt
    obj.ИмяФайла = excel_path
    obj.Записать()  # без проведения; ТЧ Расшифровка пустая, наполнит fill_rasshifrovka()
    print(f"  [CREATE] Документ А_РасшифровкаЛистов создан: №{str(obj.Номер).strip()} от {dt}")
    return obj.Ссылка


def _find_doc_by_period(conn, excel_path):
    ef = next((e for e in config.EXCEL_FILES if e["path"] == excel_path), None)
    if not ef:
        return None
    y, m, d = map(int, ef["period"].split("-"))
    dt = datetime(y, m, d, 12, 0, 0)
    q = conn.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
    ИЗ Документ.А_РасшифровкаЛистов
    ГДЕ Дата = &Дата И НЕ ПометкаУдаления"""
    q.УстановитьПараметр("Дата", dt)
    tz = q.Выполнить().Выгрузить()
    return tz.Получить(0).Ссылка if tz.Количество() > 0 else None


def backup_existing_rows(conn, doc_ref, out_dir):
    obj = doc_ref.ПолучитьОбъект()
    rows = []
    for row in obj.Расшифровка:
        ref = row.Подразделение
        is_empty = ref.Пустая() if ref else True
        rows.append({
            "ИмяЛиста": str(row.ИмяЛиста),
            "ПодразделениеСтрока": str(row.ПодразделениеСтрока),
            "Подразделение_uuid": None if is_empty else uuid_str(conn, ref),
            "Подразделение_имя": None if is_empty else str(ref.Наименование),
        })
    if not rows:
        return None
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = out_dir / f"13_backup_{str(obj.Номер).strip()}_{ts}.json"
    out.write_text(
        json.dumps(
            {"doc_number": str(obj.Номер), "backed_up_at": ts,
             "rows_count": len(rows), "rows": rows},
            ensure_ascii=False, indent=2),
        encoding="utf-8")
    return out


def fill_rasshifrovka(conn, doc_ref, excel_path, sheets, mapping, dry_run=False):
    obj = doc_ref.ПолучитьОбъект()
    rows_detail = []

    if not dry_run:
        obj.ИмяФайла = excel_path
        obj.Расшифровка.Очистить()

    for sheet_name in sheets:
        m = mapping.get(sheet_name)
        detail = {
            "sheet_name": sheet_name,
            "match_type": m["match_type"] if m else "absent",
            "struct_uuid": None,
            "struct_name": None,
            "action": "empty",
        }

        if not dry_run:
            row = obj.Расшифровка.Добавить()
            row.ИмяЛиста = sheet_name
            row.ПодразделениеСтрока = sheet_name[:150]

        dir_uuid = m.get("direction_uuid") if m else ""
        inc_children = bool(m.get("include_children", False)) if m else False

        if m and m.get("match_type") in AUTO_MATCH_TYPES and m.get("struct_uuid"):
            detail["struct_uuid"] = m["struct_uuid"]
            detail["struct_name"] = m.get("struct_name")
            detail["action"] = "filled"
            if not dry_run:
                uid = conn.NewObject("УникальныйИдентификатор", m["struct_uuid"])
                row.Подразделение = conn.Справочники.СтруктураПредприятия.ПолучитьСсылку(uid)

        detail["direction_uuid"] = dir_uuid or None
        detail["direction_name"] = (m.get("direction_name") if m else None) or None
        detail["include_children"] = inc_children

        if not dry_run and dir_uuid:
            dir_ref_id = conn.NewObject("УникальныйИдентификатор", dir_uuid)
            row.НаправлениеДеятельности = conn.Справочники.НаправленияДеятельности.ПолучитьСсылку(dir_ref_id)
        if not dry_run:
            row.ВключатьДочерние = inc_children

        rows_detail.append(detail)

    if not dry_run:
        obj.Записать()

    total = len(rows_detail)
    with_dep = sum(1 for r in rows_detail if r["action"] == "filled")
    with_dir = sum(1 for r in rows_detail if r.get("direction_uuid"))
    incch_count = sum(1 for r in rows_detail if r.get("include_children"))
    return {
        "total": total,
        "with_dep": with_dep,
        "empty_dep": total - with_dep,
        "with_direction": with_dir,
        "include_children_rows": incch_count,
        "rows": rows_detail,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="Напечатать план заполнения без модификации базы.")
    ap.add_argument("--month", default=None,
                    help="Опционально: YYYY-MM. Обрабатывать только Excel-файлы этого месяца. "
                         "Если документа А_РасшифровкаЛистов для месяца нет — создать.")
    args = ap.parse_args()

    mode = "DRY-RUN" if args.dry_run else "LIVE"
    month_filter = args.month
    if month_filter:
        print(f"=== Режим: {mode}, месяц: {month_filter} ===\n")
    else:
        print(f"=== Режим: {mode} (все месяцы из config.EXCEL_FILES) ===\n")

    conn = connect_erp()
    mapping_raw = json.loads(
        (config.JSON_DIR / "07_mapping_sheet_to_struct.json").read_text(encoding="utf-8"))
    mapping_by_sheet = {m["sheet_name"]: m for m in mapping_raw["mappings"]}
    print(f"Mappings загружено: {len(mapping_by_sheet)}")

    entries = []
    for ef in config.EXCEL_FILES:
        path, label = ef["path"], ef["label"]
        if month_filter and not ef["period"].startswith(month_filter):
            continue
        print(f"\n--- {label} ---")
        visible, hidden = read_visible_and_hidden(path)
        print(f"  Листов видимых: {len(visible)}, скрытых (пропущены): {len(hidden)}")

        doc_ref = find_doc_by_file(conn, path)
        if not doc_ref:
            if month_filter and not args.dry_run:
                doc_ref = create_doc_for_period(conn, path)
            elif month_filter and args.dry_run:
                print(f"  [DRY-RUN] Документ для '{path}' не найден — был бы создан")
                continue
            else:
                raise RuntimeError(f"Документ для '{path}' не найден (ни по basename, ни по дате)")
        print(f"  Документ: UUID={uuid_str(conn, doc_ref)}")

        backup_file = backup_existing_rows(conn, doc_ref, config.JSON_DIR)
        if backup_file:
            print(f"  Backup: {backup_file.name}")
        else:
            print(f"  Backup: ТЧ пуста, backup не требуется")

        res = fill_rasshifrovka(conn, doc_ref, path, visible, mapping_by_sheet,
                                dry_run=args.dry_run)
        print(f"  Всего: {res['total']}, с Подразделением: {res['with_dep']}, "
              f"пустых: {res['empty_dep']}, с Направлением: {res['with_direction']}, "
              f"свод-строк: {res['include_children_rows']}")

        entries.append({
            "file": label,
            "excel_path": path,
            "doc_uuid": uuid_str(conn, doc_ref),
            "total": res["total"],
            "with_dep": res["with_dep"],
            "empty_dep": res["empty_dep"],
            "with_direction": res["with_direction"],
            "include_children_rows": res["include_children_rows"],
            "rows": res["rows"],
            "skipped_hidden": hidden,
            "backup_file": backup_file.name if backup_file else None,
        })

    if not entries:
        print(f"\n[!] Не обработано ни одного файла. month_filter={month_filter!r}.")
        return
    suffix = "_dryrun" if args.dry_run else ""
    if month_filter:
        suffix += f"_{month_filter}"
    out = config.JSON_DIR / f"13_fill_rasshifrovka_log{suffix}.json"
    out.write_text(
        json.dumps({"run_at": datetime.now().isoformat(), "mode": mode,
                    "entries": entries},
                   ensure_ascii=False, indent=2),
        encoding="utf-8")
    print(f"\nЛог: {out}")


if __name__ == "__main__":
    main()
