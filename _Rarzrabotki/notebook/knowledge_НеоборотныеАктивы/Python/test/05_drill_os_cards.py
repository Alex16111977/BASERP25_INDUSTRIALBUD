# -*- coding: utf-8 -*-
"""
05_drill_os_cards.py — карточки конкретных ОС из задвоенного документа:
  - Гайковерт BOSCH
  - Вышка 1,2*2,0
  - Фильтр ECOSOFT

Цель: понять параметры амортизации (метод управ vs бух/нал, СПИ, стоимость).

Источники:
  Справочник.ОбъектыЭксплуатации       — карточка ОС
  РС.ПервоначальныеСведенияОСУпр       — параметры управ.учёта (метод, СПИ)
  РС.ПервоначальныеСведенияОСБух       — параметры бух.учёта
  РС.ПервоначальныеСведенияОСНалогОС   — параметры нал.учёта
  РС.СтоимостьОСУпр / РН.СтоимостьОС    — балансовая стоимость
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import connect_erp, save_csv, fail, money

OBJECTS = [
    "Гайковерт удар.акумулятор. BOSCH 18B GDS 18V-1050 H 2*5Ah",
    "Вишка Варіант 1,2*2,0 6+1 вис 8,7",
    "Фільтр зворотного осмосу ECOSOFT Robust mini+бак+кран д/набора води",
]


def find_osmd(erp):
    """Вернуть список реквизитов и связанных РС для ОбъектыЭксплуатации."""
    md = erp.Метаданные.Справочники.ОбъектыЭксплуатации
    rekvs = [r.Имя for r in md.Реквизиты]
    return rekvs


def find_amortization_rs(erp):
    """Найти все РС/РН, связанные с амортизацией/стоимостью ОС."""
    keywords = ["Амортизац", "СпособыНачисления", "Параметры", "Стоимость", "СрокИспользования", "ПервоначальныеСведения"]
    found = []
    for rs in erp.Метаданные.РегистрыСведений:
        name = rs.Имя
        if any(k in name and ("ОС" in name or "Эксплуат" in name) for k in keywords):
            found.append(("РС", name))
    for rn in erp.Метаданные.РегистрыНакопления:
        name = rn.Имя
        if any(k in name and ("ОС" in name or "Эксплуат" in name) for k in keywords):
            found.append(("РН", name))
    return found


def find_obj(erp, name):
    """Найти ОС по точному наименованию через запрос (НайтиПоНаименованию не работает с длинными именами)."""
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Н", name)
    q.Текст = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.ОбъектыЭксплуатации ГДЕ Наименование = &Н"
    try:
        sel = q.Выполнить().Выбрать()
        if sel.Следующий():
            return sel.Ссылка
    except Exception:
        pass
    return None


def dump_reg_for_obj(erp, reg_type, reg_name, obj_ref):
    """Выгрузить все строки РС/РН связанные с конкретным ОС, найти измерение ОбъектыЭксплуатации."""
    md = erp.Метаданные[reg_type + "ыСведений" if reg_type == "РС" else "ыНакопления"][reg_name]
    izm = [d.Имя for d in md.Измерения]
    rsc = [r.Имя for r in md.Ресурсы]
    rkv = [r.Имя for r in md.Реквизиты]

    cols = izm + rsc + rkv
    # Найти подходящее измерение для ОС
    os_dim = None
    for d in izm:
        if d in ("ОбъектУчета", "ОсновноеСредство", "ОбъектЭксплуатации", "ОбъектыЭксплуатации"):
            os_dim = d
            break
    if not os_dim:
        return None, f"нет измерения для ОС в {reg_name}"

    cols_str = ", ".join(["Т." + c + " КАК " + c for c in cols] + ["Т.Период КАК Период"] if reg_type == "РС" else ["Т." + c + " КАК " + c for c in cols])
    table = "РегистрСведений" if reg_type == "РС" else "РегистрНакопления"

    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("ОС", obj_ref)
    q.Текст = f"ВЫБРАТЬ ПЕРВЫЕ 20 {cols_str} ИЗ {table}.{reg_name} КАК Т ГДЕ Т.{os_dim} = &ОС"
    try:
        rez = q.Выполнить().Выгрузить()
    except Exception as e:
        return None, fail(e)

    rows = []
    for i in range(rez.Количество()):
        r = rez.Получить(i)
        row = {}
        for c in cols:
            try:
                v = getattr(r, c)
                row[c] = str(v) if v else ""
            except AttributeError:
                row[c] = ""
        rows.append(row)
    return rows, None


def main():
    erp = connect_erp()
    rsrn_list = find_amortization_rs(erp)
    print(f"=== Найдено связанных РС/РН: {len(rsrn_list)} ===")
    for t, n in rsrn_list:
        print(f"  {t}.{n}")

    print()
    for obj_name in OBJECTS:
        print(f"\n{'='*100}")
        print(f"=== ОС: {obj_name}")
        print(f"{'='*100}")
        obj_ref = find_obj(erp, obj_name)
        if not erp.ЗначениеЗаполнено(obj_ref):
            print(f"  НЕ НАЙДЕН в Справочник.ОбъектыЭксплуатации")
            continue

        # Главные реквизиты карточки
        try:
            obj = obj_ref.ПолучитьОбъект()
            for fld in ["Наименование", "Код", "ГруппаУчета", "ИнвентарныйНомер", "Подразделение",
                        "ОбъектПринятийКУчету"]:
                try:
                    v = getattr(obj, fld)
                    print(f"  {fld}: {v}")
                except AttributeError:
                    pass
        except Exception as e:
            print(f"  FAIL получить объект: {fail(e)}")

        # Прогон по РС/РН
        for t, n in rsrn_list:
            rows, err = dump_reg_for_obj(erp, t, n, obj_ref)
            if err:
                # тихо skip несовместимые регистры
                continue
            if not rows:
                continue
            print(f"\n  --- {t}.{n} ({len(rows)} строк) ---")
            for r in rows:
                # Выводим только непустые поля
                nz = [f"{k}={v}" for k, v in r.items() if v and v != "Пустая ссылка" and v != "0" and v != "False"]
                if nz:
                    print(f"    " + " | ".join(nz[:15]))

    return 0


if __name__ == "__main__":
    sys.exit(main())
