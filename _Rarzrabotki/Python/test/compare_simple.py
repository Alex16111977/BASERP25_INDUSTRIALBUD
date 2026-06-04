# -*- coding: utf-8 -*-
import sys, io, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

print("Старт", flush=True)

# Эталон 000Ц-000286 от 28.12.2025
qe = erp.NewObject("Запрос")
qe.УстановитьПараметр("Д1", datetime.datetime(2025, 12, 28, 0, 0, 0))
qe.УстановитьПараметр("Д2", datetime.datetime(2025, 12, 28, 23, 59, 59))
qe.УстановитьПараметр("Σ", 5194594.73)
qe.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.СписаниеБезналичныхДенежныхСредств КАК Д
ГДЕ Д.Дата МЕЖДУ &Д1 И &Д2 И Д.СуммаДокумента = &Σ
"""
re = qe.Выполнить().Выгрузить()
print(f"Эталон найден: {re.Количество()}", flush=True)

# Неисправленный 000Ц-000228 — по сумме 897.45
qn = erp.NewObject("Запрос")
qn.УстановитьПараметр("Д1", datetime.datetime(2025, 12, 16, 0, 0, 0))
qn.УстановитьПараметр("Д2", datetime.datetime(2025, 12, 16, 23, 59, 59))
qn.УстановитьПараметр("Σ", 897.45)
qn.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.СписаниеБезналичныхДенежныхСредств КАК Д
ГДЕ Д.Дата МЕЖДУ &Д1 И &Д2 И Д.СуммаДокумента = &Σ
"""
rn = qn.Выполнить().Выгрузить()
print(f"Неисправленный найден: {rn.Количество()}", flush=True)


def дамп(док, метка):
    obj = док.ПолучитьОбъект()
    print(f"\n{'='*100}", flush=True)
    print(f"{метка}: {S(док)}", flush=True)
    print(f"  Σ={obj.СуммаДокумента}", flush=True)
    fields = ["Проведен", "А_Обработан", "А_ВведенВЕРП", "А_ОбработанКазна", "А_Необновлять",
              "Контрагент", "Партнер", "Подразделение", "Договор", "ОбъектРасчетов",
              "ХозяйственнаяОперация", "БанковскийСчет", "СтатьяДвиженияДенежныхСредств",
              "НаправлениеДеятельности", "А_ОказаниеУслугМеждуПодразделениями"]
    for f in fields:
        try:
            v = getattr(obj, f, None)
            if v is None:
                s = "(None)"
            elif isinstance(v, bool):
                s = str(v)
            elif hasattr(v, 'Пустая') and v.Пустая():
                s = "(ПустаяСсылка)"
            else:
                s = str(S(v))[:55]
            print(f"  {f:<38}: {s}", flush=True)
        except Exception as e:
            print(f"  {f:<38}: <ОШИБКА: {e}>", flush=True)
    # РП
    print(f"  РасшифровкаПлатежа ({obj.РасшифровкаПлатежа.Количество()} строк):", flush=True)
    for i in range(obj.РасшифровкаПлатежа.Количество()):
        рп = obj.РасшифровкаПлатежа.Получить(i)
        for f in ["Партнер", "ДоговорСЗаказчиком", "А_Договор", "ОбъектРасчетов",
                  "Сумма", "СуммаВзаиморасчетов", "Подразделение"]:
            try:
                v = getattr(рп, f, None)
                if v is None:
                    s = "(None)"
                elif hasattr(v, 'Пустая') and v.Пустая():
                    s = "(ПустаяСсылка)"
                else:
                    s = str(S(v))[:50]
                print(f"    Стр{i+1}.{f:<32}: {s}", flush=True)
            except Exception as e:
                print(f"    Стр{i+1}.{f:<32}: <ОШИБКА>", flush=True)


if re.Количество() > 0:
    дамп(re.Получить(0).Ссылка, "ЭТАЛОН 000Ц-000286")
if rn.Количество() > 0:
    дамп(rn.Получить(0).Ссылка, "НЕИСПРАВЛ 000Ц-000228")

print("\nГОТОВО", flush=True)
