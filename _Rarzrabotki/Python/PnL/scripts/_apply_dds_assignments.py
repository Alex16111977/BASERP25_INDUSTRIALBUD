"""Этап 1.4-1.5 — применяет dds_to_pl_assignments.json в базу через COM.

Порядок:
  1. resolve_duplicates  — удалить дубли ДДС из ТЧ указанных PL
  2. add_to_tch          — добавить ДДС в ТЧ существующих PL
  3. create_new_pl       — создать новые PL-статьи с флагом А_СозданоАвтоматически=Истина
  4. fix_header_empty    — оставить пустыми где указано

ВАЖНО: выполнять ДО загрузки модуля объекта с проверкой уникальности (иначе пропуски при записи из-за временного состояния).
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, is_manually_locked, uuid_str


def find_pl_by_code(conn, code):
    ref = conn.Справочники.А_Статьи_PL.НайтиПоКоду(code)
    if ref and not ref.Пустая():
        return ref
    return None


def find_dds_by_uuid(conn, uuid):
    uid = conn.NewObject("УникальныйИдентификатор", uuid)
    return conn.Справочники.СтатьиДвиженияДенежныхСредств.ПолучитьСсылку(uid)


def find_group_by_name(conn, name):
    ref = conn.Справочники.А_ГруппаСтатей_PL.НайтиПоНаименованию(name, True)
    if ref and not ref.Пустая():
        return ref
    return None


def main():
    data = json.loads((config.JSON_DIR / "dds_to_pl_assignments.json").read_text(encoding="utf-8"))
    conn = connect_erp()

    log = {"resolve_duplicates": [], "add_to_tch": [], "create_new_pl": [], "errors": [], "locked": []}

    # ===== 1. RESOLVE DUPLICATES — убираем ДДС из указанных PL =====
    print("=== 1. Устраняю дубли ДДС ===")
    for dup in data.get("resolve_duplicates", []):
        dds_uuid = dup["dds_uuid"]
        dds_name = dup["dds_name"]
        for pl_code in dup.get("remove_from_pl_codes", []):
            pl_ref = find_pl_by_code(conn, pl_code)
            if not pl_ref:
                print(f"  SKIP: PL код {pl_code} не найден")
                continue
            pl_name = str(pl_ref)
            obj = pl_ref.ПолучитьОбъект()
            if is_manually_locked(obj):
                print(f"  LOCK {pl_name[:45]:45} (А_РучнаяКорректировка=Истина — пропущен)")
                log["locked"].append({"pl": pl_name, "stage": "resolve_duplicates"})
                continue
            # удалить строки с этой ДДС
            removed = 0
            tch = obj.Статьи
            i = tch.Количество() - 1
            while i >= 0:
                row = tch.Получить(i)
                if row.СтатьяДвиженияДенежныхСредств and uuid_str(conn, row.СтатьяДвиженияДенежныхСредств) == dds_uuid:
                    tch.Удалить(i)
                    removed += 1
                i -= 1
            # если удалили последнюю запись — добавить из шапки (ТЧ не должна быть пустой по правилу 2)
            if tch.Количество() == 0 and obj.СтатьяДвиженияДенежныхСредств and not obj.СтатьяДвиженияДенежныхСредств.Пустая():
                if uuid_str(conn, obj.СтатьяДвиженияДенежныхСредств) != dds_uuid:
                    stroka = tch.Добавить()
                    stroka.СтатьяДвиженияДенежныхСредств = obj.СтатьяДвиженияДенежныхСредств
            # если в шапке тоже эта ДДС — очистить её
            if obj.СтатьяДвиженияДенежныхСредств and uuid_str(conn, obj.СтатьяДвиженияДенежныхСредств) == dds_uuid:
                # заменить на первую из ТЧ
                if tch.Количество() > 0:
                    obj.СтатьяДвиженияДенежныхСредств = tch.Получить(0).СтатьяДвиженияДенежныхСредств
            try:
                obj.Записать()
                print(f"  {pl_name[:45]:45} -- удалено строк: {removed}")
                log["resolve_duplicates"].append({"pl": pl_name, "dds": dds_name, "removed": removed})
            except Exception as e:
                print(f"  ERROR: {pl_name} — {e}")
                log["errors"].append(f"resolve_duplicates {pl_name}: {e}")

    # ===== 2. ADD_TO_TCH — добавляем ДДС в ТЧ существующих PL =====
    print("\n=== 2. Добавляю ДДС в ТЧ существующих PL ===")
    for item in data.get("add_to_tch", []):
        dds_ref = find_dds_by_uuid(conn, item["dds_uuid"])
        pl_ref = find_pl_by_code(conn, item["target_pl_code"])
        if not pl_ref:
            print(f"  SKIP: PL код {item['target_pl_code']} не найден")
            continue
        pl_name = str(pl_ref)
        obj = pl_ref.ПолучитьОбъект()
        if is_manually_locked(obj):
            print(f"  LOCK {pl_name[:40]:40} (А_РучнаяКорректировка=Истина — пропущен)")
            log["locked"].append({"pl": pl_name, "stage": "add_to_tch"})
            continue

        # Проверить, нет ли уже такой ДДС в ТЧ
        already = False
        for i in range(obj.Статьи.Количество()):
            row = obj.Статьи.Получить(i)
            if row.СтатьяДвиженияДенежныхСредств and uuid_str(conn, row.СтатьяДвиженияДенежныхСредств) == item["dds_uuid"]:
                already = True
                break

        if already:
            print(f"  SKIP: {pl_name[:40]} уже содержит ДДС {item['dds_name']}")
            continue

        # Добавить строку
        stroka = obj.Статьи.Добавить()
        stroka.СтатьяДвиженияДенежныхСредств = dds_ref

        # Если шапочная ДДС пуста — установить туда первую
        if not obj.СтатьяДвиженияДенежныхСредств or obj.СтатьяДвиженияДенежныхСредств.Пустая():
            obj.СтатьяДвиженияДенежныхСредств = obj.Статьи.Получить(0).СтатьяДвиженияДенежныхСредств

        try:
            obj.Записать()
            print(f"  {pl_name[:40]:40} + ДДС '{item['dds_name'][:30]:30}' ({item['sum_3m']:>12,.0f} ₴)")
            log["add_to_tch"].append({"pl": pl_name, "dds": item["dds_name"]})
        except Exception as e:
            print(f"  ERROR: {pl_name} — {e}")
            log["errors"].append(f"add_to_tch {pl_name}: {e}")

    # ===== 3. CREATE_NEW_PL =====
    print("\n=== 3. Создаю новые PL-статьи ===")
    for item in data.get("create_new_pl", []):
        # проверить что не существует
        existing = conn.Справочники.А_Статьи_PL.НайтиПоНаименованию(item["new_pl_name"], True)
        if existing and not existing.Пустая():
            print(f"  SKIP: '{item['new_pl_name']}' уже существует")
            continue

        group_ref = find_group_by_name(conn, item["new_pl_group"])
        if not group_ref:
            print(f"  ERROR: группа '{item['new_pl_group']}' не найдена для '{item['new_pl_name']}'")
            continue

        dds_ref = find_dds_by_uuid(conn, item["dds_uuid"])

        obj = conn.Справочники.А_Статьи_PL.СоздатьЭлемент()
        obj.Наименование = item["new_pl_name"]
        obj.Группа = group_ref
        obj.СтатьяДвиженияДенежныхСредств = dds_ref
        # ТипСтатьи — через перечисление
        try:
            obj.ТипСтатьи = getattr(conn.Перечисления.А_ТипСтатьиPL, item["new_pl_type"])
        except Exception:
            pass
        # Маркер авто
        obj.А_СозданоАвтоматически = True

        # ТЧ Статьи — 1 строка с этой ДДС
        stroka = obj.Статьи.Добавить()
        stroka.СтатьяДвиженияДенежныхСредств = dds_ref

        try:
            obj.Записать()
            print(f"  NEW '{item['new_pl_name'][:45]:45}' group={item['new_pl_group'][:20]:20} ({item['sum_3m']:>12,.0f} ₴)")
            log["create_new_pl"].append({"name": item["new_pl_name"], "code": str(obj.Код)})
        except Exception as e:
            print(f"  ERROR: {item['new_pl_name']} — {e}")
            log["errors"].append(f"create_new_pl {item['new_pl_name']}: {e}")

    # Сохранить лог
    log_path = config.JSON_DIR / "dds_applied_log.json"
    log_path.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\n=== ИТОГО ===")
    print(f"  Дубли устранены: {len(log['resolve_duplicates'])}")
    print(f"  Добавлено в ТЧ:  {len(log['add_to_tch'])}")
    print(f"  Новых PL:        {len(log['create_new_pl'])}")
    print(f"  Защищено (lock): {len(log['locked'])}")
    print(f"  Ошибок:          {len(log['errors'])}")
    print(f"Лог: {log_path}")


if __name__ == "__main__":
    main()
