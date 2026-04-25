"""Проверка: вариант А_ОтчетPL_РасшифровкаПоНаправлению и ПользовательскиеНастройки."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp


def main():
    conn = connect_erp()
    # Load variant
    report = conn.Отчеты.А_ОтчетPL.Создать()
    schema = report.СхемаКомпоновкиДанных
    variant = schema.ВариантыНастроек.Найти("А_ОтчетPL_РасшифровкаПоНаправлению")
    if variant is None:
        print("ERR: variant not found")
        return
    print("Variant found")

    composer = conn.NewObject("КомпоновщикНастроекКомпоновкиДанных")
    source = conn.NewObject("ИсточникДоступныхНастроекКомпоновкиДанных", schema)
    composer.Инициализировать(source)
    composer.ЗагрузитьНастройки(variant.Настройки)

    # --- Variant settings: Filter items ---
    print("\n== Variant settings - Filter items ==")
    for i in range(composer.Настройки.Отбор.Элементы.Количество()):
        e = composer.Настройки.Отбор.Элементы.Получить(i)
        left = "?"
        try:
            left = str(e.ЛевоеЗначение)
        except Exception:
            pass
        right = "?"
        try:
            right = str(e.ПравоеЗначение)
        except Exception:
            pass
        use = "?"
        try:
            use = bool(e.Использование)
        except Exception:
            pass
        view_mode = "?"
        try:
            view_mode = str(e.РежимОтображения)
        except Exception:
            pass
        user_id = "?"
        try:
            user_id = str(e.ИдентификаторПользовательскойНастройки)
        except Exception:
            pass
        print(f"  [{i}] use={use} left={left!r} right={right!r} view={view_mode} userID={user_id}")

    # --- User settings: all items ---
    print("\n== User settings - all items ==")
    user_items = composer.ПользовательскиеНастройки.Элементы
    for i in range(user_items.Количество()):
        e = user_items.Получить(i)
        # Type
        tp = type(e).__name__
        # Common attributes
        info = {}
        for attr in ("Параметр", "ЛевоеЗначение", "ПравоеЗначение", "Значение",
                     "Использование", "ИдентификаторПользовательскойНастройки"):
            try:
                info[attr] = str(getattr(e, attr))
            except Exception:
                pass
        print(f"  [{i}] {tp}  {info}")


if __name__ == "__main__":
    main()
