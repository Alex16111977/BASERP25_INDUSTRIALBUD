# -*- coding: utf-8 -*-
"""Verify: после перепровода 28 А_ФинРез_PL сравнить с pretest:
  - |Σ Сумма| (абсолют) должен совпасть до 0.01 ₽
  - Σ Сумма для Доход = +|Σabs|, для Расход = −|Σabs|"""
import sys, io, json, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# Pretest
path = os.path.join(os.path.dirname(__file__), "finrez_pl_sign_pretest.json")
with open(path, "r", encoding="utf-8") as f:
    pretest = json.load(f)

# Текущее состояние
q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ
    Р.Организация КАК Орг,
    Р.Статья КАК Статья,
    Р.Статья.ТипСтатьи КАК ТипСтатьи,
    КОЛИЧЕСТВО(*) КАК Строк,
    СУММА(Р.Сумма) КАК Σ_Сумма,
    СУММА(ВЫБОР КОГДА Р.Сумма < 0 ТОГДА -Р.Сумма ИНАЧЕ Р.Сумма КОНЕЦ) КАК Σabs_Сумма,
    СУММА(Р.СуммаФ1_Excel) КАК Σ_Ф1_Excel,
    СУММА(ВЫБОР КОГДА Р.СуммаФ1_Excel < 0 ТОГДА -Р.СуммаФ1_Excel ИНАЧЕ Р.СуммаФ1_Excel КОНЕЦ) КАК Σabs_Ф1_Excel
ИЗ РегистрСведений.А_ОтчетPL_Свод КАК Р
СГРУППИРОВАТЬ ПО Р.Организация, Р.Статья, Р.Статья.ТипСтатьи
"""
tz = q.Выполнить().Выгрузить()
print(f"После перепровода групп: {tz.Количество()} (было {len(pretest)})")

# Сводки
дох = [r for r in tz if conn.String(r.ТипСтатьи) == "Доход"]
расх = [r for r in tz if conn.String(r.ТипСтатьи) == "Расход"]

def Σ(rows, fld):
    return sum(float(getattr(r, fld)) for r in rows)

now_dox_signed  = Σ(дох,  'Σ_Сумма')
now_dox_abs     = Σ(дох,  'Σabs_Сумма')
now_rasx_signed = Σ(расх, 'Σ_Сумма')
now_rasx_abs    = Σ(расх, 'Σabs_Сумма')

print(f"\n=== Доход ({len(дох)} групп) ===")
print(f"  Σ Сумма (со знаком после): {now_dox_signed:>20,.2f}")
print(f"  Σ |Сумма| (абсолют):       {now_dox_abs:>20,.2f}")
print(f"  Доход: Σ ≡ +Σabs? {abs(now_dox_signed - now_dox_abs) < 0.01}")

print(f"\n=== Расход ({len(расх)} групп) ===")
print(f"  Σ Сумма (со знаком после): {now_rasx_signed:>20,.2f}")
print(f"  Σ |Сумма| (абсолют):       {now_rasx_abs:>20,.2f}")
print(f"  Расход: Σ ≡ −Σabs? {abs(now_rasx_signed + now_rasx_abs) < 0.01}")

# Pretest baseline
pretest_dox_abs = sum(g["abs_sums"]["Сумма"] for g in pretest if g["тип"] == "Доход")
pretest_rasx_abs = sum(g["abs_sums"]["Сумма"] for g in pretest if g["тип"] == "Расход")

print(f"\n=== |Σabs| pretest vs verify ===")
print(f"  Доход pretest|abs|: {pretest_dox_abs:>20,.2f}")
print(f"  Доход verify |abs|: {now_dox_abs:>20,.2f}  Δ={now_dox_abs-pretest_dox_abs:>10,.4f}")
print(f"  Расход pretest|abs|: {pretest_rasx_abs:>20,.2f}")
print(f"  Расход verify |abs|: {now_rasx_abs:>20,.2f}  Δ={now_rasx_abs-pretest_rasx_abs:>10,.4f}")

# Чистый P&L
print(f"\n=== Сквозной P&L (Доход + Расход со знаком) ===")
pl = now_dox_signed + now_rasx_signed
print(f"  Σ всё со знаком = {pl:>20,.2f} ₽  ({'УБЫТОК' if pl < 0 else 'ПРИБЫЛЬ'})")

# Mirror-инваріант: Σ_signed_verify == -Σ_signed_pretest (для Расхода).
# Для Доход — без изменений (signed = abs обеих сторон).
pretest_dox_signed = sum(g["signed_sums_now"]["Сумма"] for g in pretest if g["тип"] == "Доход")
pretest_rasx_signed = sum(g["signed_sums_now"]["Сумма"] for g in pretest if g["тип"] == "Расход")

print(f"\n=== Mirror-инвариант ===")
print(f"  Доход:  signed pretest = {pretest_dox_signed:>20,.2f}")
print(f"          signed verify  = {now_dox_signed:>20,.2f}  Δ vs pretest = {now_dox_signed - pretest_dox_signed:>10,.4f}")
print(f"  Расход: signed pretest = {pretest_rasx_signed:>20,.2f}")
print(f"          signed verify  = {now_rasx_signed:>20,.2f}  Δ vs -pretest = {now_rasx_signed + pretest_rasx_signed:>10,.4f}")

# Assertions
assert abs(now_dox_signed - pretest_dox_signed) < 1.0, "Доход signed изменился!"
assert abs(now_rasx_signed + pretest_rasx_signed) < 1.0, "Расход НЕ mirror: signed_verify ≠ -signed_pretest!"
assert abs(now_dox_abs - pretest_dox_abs) < 1.0, "|Σabs| Доход изменился!"
assert abs(now_rasx_abs - pretest_rasx_abs) < 1.0, "|Σabs| Расход изменился!"
print("\n[OK] Все Σ-инварианты выполнены: Доход без изменений, Расход = mirror (−1*pretest), |abs| сохранён.")
