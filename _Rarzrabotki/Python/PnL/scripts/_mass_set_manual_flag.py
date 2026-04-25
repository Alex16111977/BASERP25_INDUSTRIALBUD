"""Разово проставить А_РучнаяКорректировка=Истина всем существующим элементам Справочник.А_Статьи_PL.

Запуск после развёртывания реквизита в конфигурации (см. план defs-plan).
Обходит ПередЗаписью через ОбменДанными.Загрузка=Истина — чтобы целостностные проверки
модуля объекта (уникальность ДДС, обязательная шапочная ДДС для расходов) не срывали массовую
простановку флага. Запись самого флага бизнес-данные не меняет.

Usage:
    python scripts/_mass_set_manual_flag.py --dry-run   # только посчитать
    python scripts/_mass_set_manual_flag.py              # боевой прогон
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp, run_query


QUERY_CANDIDATES = """
ВЫБРАТЬ
    PL.Ссылка КАК Ссылка,
    PL.Наименование КАК Наименование,
    PL.Код КАК Код,
    PL.А_РучнаяКорректировка КАК Флаг
ИЗ Справочник.А_Статьи_PL КАК PL
ГДЕ НЕ PL.ЭтоГруппа
    И НЕ PL.ПометкаУдаления
УПОРЯДОЧИТЬ ПО PL.Код
"""


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true", help="только показать счётчик, ничего не писать")
    args = parser.parse_args()

    conn = connect_erp()
    tz = run_query(conn, QUERY_CANDIDATES)

    total = tz.Количество()
    already = 0
    to_update = []
    for i in range(total):
        row = tz.Получить(i)
        if bool(row.Флаг):
            already += 1
        else:
            to_update.append((row.Ссылка, str(row.Наименование), str(row.Код)))

    print(f"Всего не-групп и не-помеченных-на-удаление статей: {total}")
    print(f"Уже защищено (А_РучнаяКорректировка=Истина): {already}")
    print(f"Будет проставлено: {len(to_update)}")

    if args.dry_run:
        print("\n[DRY-RUN] Ничего не записываю.")
        for ref, name, code in to_update[:20]:
            print(f"  {code} {name}")
        if len(to_update) > 20:
            print(f"  ... и ещё {len(to_update) - 20}")
        return

    if not to_update:
        print("\nНечего делать.")
        return

    print("\n=== Простановка флага ===")
    ok = 0
    errors = []
    for ref, name, code in to_update:
        try:
            obj = ref.ПолучитьОбъект()
            # Обходим ПередЗаписью (правила целостности) — запись только самого флага
            obj.ОбменДанными.Загрузка = True
            obj.А_РучнаяКорректировка = True
            obj.Записать()
            ok += 1
            print(f"  OK  {code} {name}")
        except Exception as e:
            errors.append((code, name, str(e)))
            print(f"  ERR {code} {name}: {e}")

    print(f"\n=== ИТОГО ===")
    print(f"  Проставлено: {ok}")
    print(f"  Ошибок:      {len(errors)}")
    if errors:
        print("\nОшибки:")
        for code, name, e in errors:
            print(f"  {code} {name}: {e}")


if __name__ == "__main__":
    main()
