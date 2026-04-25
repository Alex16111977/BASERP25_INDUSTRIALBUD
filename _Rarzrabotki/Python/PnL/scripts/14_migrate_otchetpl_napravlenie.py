"""Step 14 (one-off): Enrich existing Документ.А_ОтчетPL with НаправлениеДеятельности + ВключатьДочерние.

Стратегия:
1. Взять все непомеченные документы А_ОтчетPL.
2. Для каждого:
   - Посмотреть 07_mapping по ПодразделениеСтрока: если найдено, использовать direction_uuid + include_children оттуда.
   - Иначе: direction_uuid = Подразделение.А_НаправлениеДеятельности (если ссылка заполнена), include_children = False.
   - Если уже заполнено и значения совпадают — skip (idempotent).
3. Записать документ БЕЗ проведения (doc.Записать() — реквизиты шапки, не движения).
4. Лог в data/json/14_migrate_otchetpl_log.json.

Флаги:
- --dry-run: не писать, только печатать план.
- --limit N: обработать только первые N.
"""
import argparse
import json
import sys
import traceback
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str


def load_mapping_by_podr_stroka():
    raw = json.loads(
        (config.JSON_DIR / "07_mapping_sheet_to_struct.json").read_text(encoding="utf-8")
    )
    out = {}
    for m in raw["mappings"]:
        out[m["sheet_name"]] = {
            "direction_uuid": m.get("direction_uuid") or "",
            "direction_name": m.get("direction_name") or "",
            "include_children": bool(m.get("include_children", False)),
        }
    return out


def fetch_all_otchetpl_refs(conn):
    q = conn.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ Ссылка
    ИЗ Документ.А_ОтчетPL
    ГДЕ НЕ ПометкаУдаления
    УПОРЯДОЧИТЬ ПО Дата
    """
    tz = q.Выполнить().Выгрузить()
    refs = []
    for i in range(tz.Количество()):
        refs.append(tz.Получить(i).Ссылка)
    return refs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--limit", type=int, default=None)
    args = ap.parse_args()

    mode = "DRY-RUN" if args.dry_run else "LIVE"
    print(f"=== 14 Migrate А_ОтчетPL направление ({mode}) ===\n")

    conn = connect_erp()
    mapping = load_mapping_by_podr_stroka()
    print(f"Mapping: {len(mapping)} листов")

    refs = fetch_all_otchetpl_refs(conn)
    if args.limit:
        refs = refs[: args.limit]
    print(f"Документов к обработке: {len(refs)}\n")

    empty_napr_ref = conn.Справочники.НаправленияДеятельности.ПустаяСсылка()

    log = []
    updated = skipped = failed = 0

    for idx, ref in enumerate(refs, 1):
        num = "?"
        try:
            obj = ref.ПолучитьОбъект()
            num = str(obj.Номер).strip()
            date_str = str(obj.Дата)[:10]
            podr_str = str(obj.ПодразделениеСтрока) if obj.ПодразделениеСтрока else ""

            # 1) Целевые значения
            m = mapping.get(podr_str)
            target_dir_uuid = ""
            target_dir_name = ""
            target_incch = False
            source = ""

            if m and m.get("direction_uuid"):
                target_dir_uuid = m["direction_uuid"]
                target_dir_name = m["direction_name"]
                target_incch = m["include_children"]
                source = "mapping"
            elif obj.Подразделение and not obj.Подразделение.Пустая():
                napr = obj.Подразделение.А_НаправлениеДеятельности
                if napr and not napr.Пустая():
                    target_dir_uuid = uuid_str(conn, napr)
                    target_dir_name = str(napr.Наименование)
                    source = "podrazdelenie"

            # 2) Текущее состояние
            cur_napr = obj.НаправлениеДеятельности
            cur_napr_uuid = "" if (cur_napr is None or cur_napr.Пустая()) else uuid_str(conn, cur_napr)
            cur_incch = bool(obj.ВключатьДочерние)

            # 3) Нужно ли обновлять?
            need_dir = target_dir_uuid and target_dir_uuid.lower() != cur_napr_uuid.lower()
            need_incch = target_incch != cur_incch

            if not need_dir and not need_incch:
                skipped += 1
                log.append({
                    "num": num, "date": date_str, "podr_stroka": podr_str,
                    "source": source or "current_ok",
                    "action": "skip",
                    "current_direction": cur_napr_uuid,
                    "current_include_children": cur_incch,
                    "target_direction": target_dir_uuid,
                    "target_include_children": target_incch,
                })
                print(f"  [{idx:3}/{len(refs)}] SKIP {num} / {podr_str}  (napr={cur_napr_uuid or '-'} incch={cur_incch})")
                continue

            if not args.dry_run:
                if target_dir_uuid:
                    uid = conn.NewObject("УникальныйИдентификатор", target_dir_uuid)
                    obj.НаправлениеДеятельности = conn.Справочники.НаправленияДеятельности.ПолучитьСсылку(uid)
                else:
                    obj.НаправлениеДеятельности = empty_napr_ref
                obj.ВключатьДочерние = target_incch
                obj.ОбменДанными.Загрузка = True  # пропустить валидацию ПередЗаписью
                obj.Записать()

            updated += 1
            log.append({
                "num": num, "date": date_str, "podr_stroka": podr_str,
                "source": source,
                "action": "update" if not args.dry_run else "update-dry",
                "current_direction": cur_napr_uuid,
                "current_include_children": cur_incch,
                "target_direction": target_dir_uuid,
                "target_direction_name": target_dir_name,
                "target_include_children": target_incch,
            })
            tag = "DRY" if args.dry_run else "UPD"
            print(f"  [{idx:3}/{len(refs)}] {tag} {num} / {podr_str}  "
                  f"napr: {cur_napr_uuid or '-'} -> {target_dir_uuid or '-'}  incch: {cur_incch} -> {target_incch}  ({source})")

        except Exception as ex:
            failed += 1
            err = traceback.format_exc(limit=2)
            log.append({
                "num": num,
                "action": "error",
                "error": str(ex),
                "trace": err[:500],
            })
            print(f"  [{idx:3}/{len(refs)}] ERR {ex}")

    suffix = "_dryrun" if args.dry_run else ""
    out = config.JSON_DIR / f"14_migrate_otchetpl_log{suffix}.json"
    out.write_text(
        json.dumps({
            "run_at": datetime.now().isoformat(timespec="seconds"),
            "mode": mode,
            "total": len(refs),
            "updated": updated,
            "skipped": skipped,
            "failed": failed,
            "entries": log,
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"\nLog: {out}  (updated={updated}, skipped={skipped}, failed={failed})")


if __name__ == "__main__":
    main()
