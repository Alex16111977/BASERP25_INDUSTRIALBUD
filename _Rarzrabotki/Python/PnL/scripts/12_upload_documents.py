"""Step 12: upsert Документ.А_ОтчетPL by (Дата, ПодразделениеСтрока) + проведение.

- Ищет существующий документ по дате (в пределах месяца) + имени листа (ПодразделениеСтрока).
- Если найден — обновляет (очищает и заполняет заново ТЧ).
- Если нет — создаёт новый.
- В конце: Записать с проведением (РежимЗаписиДокумента.Проведение).
"""
import argparse
import json
import sys
import traceback
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str as _uuid_str


def ref_by_uuid(conn, col_name, uuid_value):
    mgr = getattr(conn.Справочники, col_name)
    uid = conn.NewObject("УникальныйИдентификатор", uuid_value)
    return mgr.ПолучитьСсылку(uid)


def find_org(conn, name):
    ref = conn.Справочники.Организации.НайтиПоНаименованию(name, True)
    if ref and not ref.Пустая():
        return ref
    return None


def find_existing_doc(conn, date_obj, podrazdelenie_stroka):
    """Найти непомеченный документ А_ОтчетPL за месяц по ПодразделениеСтрока."""
    month_start = date_obj.replace(day=1, hour=0, minute=0, second=0)
    next_month = (month_start + timedelta(days=32)).replace(day=1)
    month_end = next_month - timedelta(seconds=1)

    q = conn.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.А_ОтчетPL "
        "ГДЕ Дата МЕЖДУ &С И &ПО "
        "И ПодразделениеСтрока = &Стр И НЕ ПометкаУдаления"
    )
    q.УстановитьПараметр("С", month_start)
    q.УстановитьПараметр("ПО", month_end)
    q.УстановитьПараметр("Стр", podrazdelenie_stroka)
    tz = q.Выполнить().Выгрузить()
    if tz.Количество() > 0:
        return tz.Получить(0).Ссылка
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--month", default=None,
                    help="Опционально: YYYY-MM. Загружать только документы с date, начинающейся с указанного префикса.")
    args = ap.parse_args()

    src = config.JSON_DIR / "08_documents_to_import.json"
    data = json.loads(src.read_text(encoding="utf-8"))
    docs = data["documents"]
    if args.month:
        before = len(docs)
        docs = [d for d in docs if d.get("date", "").startswith(args.month)]
        print(f"[FILTER --month={args.month}] {before} -> {len(docs)} документов")
    if args.limit:
        docs = docs[: args.limit]

    arts = {
        a["name"]: a
        for a in json.loads(
            (config.JSON_DIR / "02_unique_articles.json").read_text(encoding="utf-8")
        )["articles"]
    }

    conn = connect_erp()
    org = find_org(conn, config.ORGANIZATION_NAME)
    if not org:
        print(f"ERROR: Организация '{config.ORGANIZATION_NAME}' not found")
        sys.exit(1)

    mode_provesti = conn.РежимЗаписиДокумента.Проведение

    log = []
    success_new = success_upd = failed = 0
    stage = "init"

    for idx, d in enumerate(docs, 1):
        sn = d["sheet_name"]
        date_str = d["date"]
        date_obj = datetime.strptime(date_str, "%Y-%m-%d").replace(
            hour=23, minute=59, second=59
        )
        podr_stroka = d["podrazdelenie_stroka"][:150]

        try:
            if args.dry_run:
                print(
                    f"  [{idx:3}/{len(docs)}] DRY {d['period_label']} / {sn}  "
                    f"rows={len(d['rows'])}"
                )
                continue

            stage = "FindExisting"
            existing_ref = find_existing_doc(conn, date_obj, podr_stroka)
            if existing_ref is not None:
                stage = "GetObject"
                doc = existing_ref.ПолучитьОбъект()
                action = "UPD"
            else:
                stage = "CreateDoc"
                doc = conn.Документы.А_ОтчетPL.СоздатьДокумент()
                doc.Дата = date_obj
                action = "NEW"

            stage = "SetHeader"
            doc.Организация = org
            doc.ПодразделениеСтрока = podr_stroka
            if d.get("podrazdelenie_uuid"):
                doc.Подразделение = ref_by_uuid(
                    conn, "СтруктураПредприятия", d["podrazdelenie_uuid"]
                )
            if d.get("direction_uuid"):
                doc.НаправлениеДеятельности = ref_by_uuid(
                    conn, "НаправленияДеятельности", d["direction_uuid"]
                )
            doc.ВключатьДочерние = bool(d.get("include_children", False))

            stage = "ClearRows"
            doc.ДанныеОтчета.Очистить()
            doc.ИтогиПоГруппам.Очистить()
            doc.ИтогиОбщие.Очистить()

            stage = "FillRows"
            # пересчёт по группам для ТЧ ИтогиПоГруппам
            sum_by_group = {}
            filled = 0
            for r in d["rows"]:
                art = arts.get(r["article_name"])
                if not art or not art.get("uuid"):
                    continue
                row = doc.ДанныеОтчета.Добавить()
                row.Статья = ref_by_uuid(conn, "А_Статьи_PL", art["uuid"])
                row.Комментарий = (r.get("comment") or "")[:1000]
                row.СуммаФома1 = float(r.get("sum_f1", 0) or 0)
                row.СуммаФорма2 = float(r.get("sum_f2", 0) or 0)
                row.Сумма = float(r.get("total", 0) or 0)
                # новые поля
                row.СтрокаЛиста = int(r.get("row_excel") or 0)
                t = r.get("type_statii") or "Расход"
                try:
                    row.ТипСтроки = getattr(conn.Перечисления.А_ТипСтатьиPL, t)
                except Exception:
                    pass
                # накопление по группе
                g = r.get("group_name") or ""
                if g:
                    sum_by_group[g] = sum_by_group.get(g, 0) + float(r.get("total", 0) or 0)
                filled += 1

            stage = "FillGroupTotals"
            for gt in d.get("group_totals", []):
                trow = doc.ИтогиПоГруппам.Добавить()
                if gt.get("group_uuid"):
                    trow.Группа = ref_by_uuid(conn, "А_ГруппаСтатей_PL", gt["group_uuid"])
                trow.СуммаФ1_Excel = float(gt.get("sum_f1_excel", 0) or 0)
                trow.СуммаФ2_Excel = float(gt.get("sum_f2_excel", 0) or 0)
                trow.Сумма_Excel = float(gt.get("sum_excel", 0) or 0)
                calc = float(sum_by_group.get(gt.get("group_name", ""), 0))
                trow.Сумма_Расчёт = calc
                trow.Расхождение = trow.Сумма_Excel - calc

            stage = "FillGeneralTotals"
            for gt in d.get("general_totals", []):
                trow = doc.ИтогиОбщие.Добавить()
                trow.Показатель = (gt.get("show") or "")[:100]
                vid = gt.get("vid")
                if vid:
                    try:
                        trow.ВидПоказателя = getattr(conn.Перечисления.А_ВидПоказателяPL, vid)
                    except Exception:
                        pass
                trow.СуммаФ1 = float(gt.get("sum_f1", 0) or 0)
                trow.СуммаФ2 = float(gt.get("sum_f2", 0) or 0)
                trow.СуммаИтого = float(gt.get("sum_total", 0) or 0)

            stage = "WritePost"
            doc.Записать(mode_provesti)

            uuid = _uuid_str(conn, doc.Ссылка)
            log.append(
                {
                    "sheet": sn,
                    "period": d["period_label"],
                    "status": "ok",
                    "action": action,
                    "uuid": uuid,
                    "rows": filled,
                    "proveden": bool(doc.Проведен),
                }
            )
            print(
                f"  [{idx:3}/{len(docs)}] {action} Проведен={bool(doc.Проведен)}  "
                f"{d['period_label']} / {sn}  rows={filled}  {uuid}"
            )
            if action == "NEW":
                success_new += 1
            else:
                success_upd += 1
        except Exception as ex:
            err = traceback.format_exc(limit=2)
            log.append(
                {
                    "sheet": sn,
                    "period": d["period_label"],
                    "status": "error",
                    "stage": stage,
                    "error": str(ex),
                    "trace": err[:500],
                }
            )
            print(
                f"  [{idx:3}/{len(docs)}] ERR [{stage}] {d['period_label']} / {sn}: {ex}"
            )
            failed += 1

    out = config.JSON_DIR / "09_import_log.json"
    out.write_text(
        json.dumps(
            {
                "run_at": datetime.now().isoformat(timespec="seconds"),
                "success_new": success_new,
                "success_upd": success_upd,
                "failed": failed,
                "entries": log,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(
        f"\nLog: {out}  (new={success_new}, upd={success_upd}, failed={failed})"
    )


if __name__ == "__main__":
    main()
