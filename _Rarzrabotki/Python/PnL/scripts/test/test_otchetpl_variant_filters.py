"""Проверка через COM: что содержит вариант А_ОтчетPL_РасшифровкаПоНаправлению
(какие фильтры/параметры) и что даёт загрузка настроек."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp


def dump_settings(conn, settings, tag):
    print(f"\n=== {tag} ===")
    # Отборы
    print("-- Отбор --")
    for i, elem in enumerate(settings.Отбор.Элементы):
        try:
            use = bool(elem.Использование)
        except Exception:
            use = "?"
        try:
            left = str(elem.ЛевоеЗначение)
        except Exception:
            left = "?"
        try:
            right = str(elem.ПравоеЗначение)
        except Exception:
            right = "?"
        print(f"  [{i}] use={use}  {left} = '{right}'")
    # Параметры
    print("-- Параметры данных --")
    for i in range(settings.ПараметрыДанных.Элементы.Количество()):
        e = settings.ПараметрыДанных.Элементы.Получить(i)
        try:
            param = str(e.Параметр)
        except Exception:
            param = "?"
        try:
            val = str(e.Значение)
        except Exception:
            val = "?"
        try:
            use = bool(e.Использование)
        except Exception:
            use = "?"
        try:
            user_id = str(e.ИдентификаторПользовательскойНастройки)
        except Exception:
            user_id = "?"
        print(f"  [{i}] use={use} {param} = {val} (userID={user_id})")


def main():
    conn = connect_erp()

    # 1) Создать отчёт и получить вариант по имени
    report = conn.Отчеты.А_ОтчетPL.Создать()
    schema = report.СхемаКомпоновкиДанных
    variant_name = "А_ОтчетPL_РасшифровкаПоНаправлению"
    variant = schema.ВариантыНастроек.Найти(variant_name)
    if variant is None:
        print(f"ERR: Вариант {variant_name} не найден!")
        return
    print(f"OK: вариант {variant_name} найден")

    # 2) Загрузить настройки в компоновщик
    composer = conn.NewObject("КомпоновщикНастроекКомпоновкиДанных")
    source = conn.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema)
    composer.Инициализировать(source)
    composer.ЗагрузитьНастройки(variant.Настройки)

    # 3) Вывести что в настройках варианта
    dump_settings(conn, composer.Настройки, "После ЗагрузитьНастройки (вариант)")
    dump_settings(conn, composer.ПользовательскиеНастройки, "ПользовательскиеНастройки")

    # 4) Поиск направления Логистика
    napr_ref = conn.Справочники.НаправленияДеятельности.НайтиПоНаименованию("Логистика")
    print(f"\nНаправление Логистика найдено: {not napr_ref.Пустая()}")

    # 5) Эмулировать установку значения фильтра НаправлениеДеятельности в пользовательских настройках
    found_napr_filter = False
    for elem in composer.ПользовательскиеНастройки.Элементы:
        type_str = str(conn.XMLСтрока(elem.Тип() if callable(getattr(elem, 'Тип', None)) else type(elem).__name__))
        # Это может быть фильтр
        try:
            left = str(elem.ЛевоеЗначение)
            if left == "НаправлениеДеятельности":
                elem.ПравоеЗначение = napr_ref
                elem.Использование = True
                found_napr_filter = True
                print(f"  Фильтр НаправлениеДеятельности установлен в Логистика")
        except Exception:
            pass

    if not found_napr_filter:
        print("  !!! Фильтр НаправлениеДеятельности не найден в ПользовательскиеНастройки")

    print("\n=== После установки фильтра ===")
    dump_settings(conn, composer.ПользовательскиеНастройки, "ПользовательскиеНастройки (after)")


if __name__ == "__main__":
    main()
