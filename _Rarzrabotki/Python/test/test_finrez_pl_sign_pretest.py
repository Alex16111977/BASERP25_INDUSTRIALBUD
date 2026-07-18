# -*- coding: utf-8 -*-
"""Pretest baseline для А_ФинРез_PL → А_ОтчетPL_Свод.
Снимает Σ |Сумма| по (Орг, Статья, ТипСтатьи) до правки знака."""
import sys, io, json, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ
    Р.Организация КАК Орг,
    Р.Статья КАК Статья,
    Р.Статья.ТипСтатьи КАК ТипСтатьи,
    КОЛИЧЕСТВО(*) КАК Строк,
    СУММА(Р.СуммаФ1_Excel) КАК Σ_Ф1_Excel,
    СУММА(Р.СуммаФ2_Excel) КАК Σ_Ф2_Excel,
    СУММА(Р.Сумма_Excel) КАК Σ_Excel,
    СУММА(Р.СуммаФ1) КАК Σ_Ф1,
    СУММА(Р.СуммаФ2) КАК Σ_Ф2,
    СУММА(Р.Сумма) КАК Σ_Сумма,
    СУММА(ВЫБОР КОГДА Р.СуммаФ1_Excel < 0 ТОГДА -Р.СуммаФ1_Excel ИНАЧЕ Р.СуммаФ1_Excel КОНЕЦ) КАК Σabs_Ф1_Excel,
    СУММА(ВЫБОР КОГДА Р.СуммаФ2_Excel < 0 ТОГДА -Р.СуммаФ2_Excel ИНАЧЕ Р.СуммаФ2_Excel КОНЕЦ) КАК Σabs_Ф2_Excel,
    СУММА(ВЫБОР КОГДА Р.Сумма_Excel < 0 ТОГДА -Р.Сумма_Excel ИНАЧЕ Р.Сумма_Excel КОНЕЦ) КАК Σabs_Excel,
    СУММА(ВЫБОР КОГДА Р.СуммаФ1 < 0 ТОГДА -Р.СуммаФ1 ИНАЧЕ Р.СуммаФ1 КОНЕЦ) КАК Σabs_Ф1,
    СУММА(ВЫБОР КОГДА Р.СуммаФ2 < 0 ТОГДА -Р.СуммаФ2 ИНАЧЕ Р.СуммаФ2 КОНЕЦ) КАК Σabs_Ф2,
    СУММА(ВЫБОР КОГДА Р.Сумма < 0 ТОГДА -Р.Сумма ИНАЧЕ Р.Сумма КОНЕЦ) КАК Σabs_Сумма
ИЗ РегистрСведений.А_ОтчетPL_Свод КАК Р
СГРУППИРОВАТЬ ПО Р.Организация, Р.Статья, Р.Статья.ТипСтатьи
"""
tz = q.Выполнить().Выгрузить()

print(f"Всего групп (Орг,Статья,ТипСтатьи): {tz.Количество()}")

# Сводка по типам
дох = [r for r in tz if conn.String(r.ТипСтатьи) == "Доход"]
расх = [r for r in tz if conn.String(r.ТипСтатьи) == "Расход"]
пуст = [r for r in tz if not (r.ТипСтатьи and conn.ЗначениеЗаполнено(r.ТипСтатьи))]

def Σ(rows, fld):
    return sum(float(getattr(r, fld)) for r in rows)

print(f"\n=== Доход ({len(дох)} групп) ===")
print(f"  Σ Сумма (со знаком как сейчас): {Σ(дох,'Σ_Сумма'):>20,.2f}")
print(f"  Σ |Сумма| (абсолют):            {Σ(дох,'Σabs_Сумма'):>20,.2f}")

print(f"\n=== Расход ({len(расх)} групп) ===")
print(f"  Σ Сумма (со знаком как сейчас): {Σ(расх,'Σ_Сумма'):>20,.2f}")
print(f"  Σ |Сумма| (абсолют):            {Σ(расх,'Σabs_Сумма'):>20,.2f}")

print(f"\n=== Без ТипСтатьи / Пустая ({len(пуст)} групп) ===")
print(f"  Σ Сумма:                         {Σ(пуст,'Σ_Сумма'):>20,.2f}")

# Сохранить baseline
out = []
for r in tz:
    out.append({
        "org": conn.String(r.Орг),
        "статья": conn.String(r.Статья),
        "тип": conn.String(r.ТипСтатьи) if (r.ТипСтатьи and conn.ЗначениеЗаполнено(r.ТипСтатьи)) else "",
        "строк": int(r.Строк),
        "abs_sums": {
            "Ф1_Excel": float(r.Σabs_Ф1_Excel),
            "Ф2_Excel": float(r.Σabs_Ф2_Excel),
            "Excel":    float(r.Σabs_Excel),
            "Ф1":       float(r.Σabs_Ф1),
            "Ф2":       float(r.Σabs_Ф2),
            "Сумма":    float(r.Σabs_Сумма),
        },
        "signed_sums_now": {
            "Сумма": float(r.Σ_Сумма),
        }
    })

path = os.path.join(os.path.dirname(__file__), "finrez_pl_sign_pretest.json")
with open(path, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
print(f"\nSaved: {path}")
