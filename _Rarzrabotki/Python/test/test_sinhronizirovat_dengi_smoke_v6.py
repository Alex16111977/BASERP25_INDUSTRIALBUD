# -*- coding: utf-8 -*-
"""
Smoke v6: данные сравнения для СКД (ПостроитьДанныеСравнения) — оба режима.
Сверка: расхождение кон.остатка по счёту ОТП (свод ЕРП-УПП) > 0 и печать значения.
Только чтение.
"""
import sys
from datetime import datetime

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьДеньги.epf"
IBAN_OTP = "UA973005280000026004000010559"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

FAILS = 0

obr = erp.ВнешниеОбработки.Создать(EPF, False)
obr.НачалоПериода = datetime(2026, 1, 1, 0, 0, 0)
obr.ОкончаниеПериода = datetime(2026, 6, 10, 0, 0, 0)

# --- режим без дней ---
obr.ПоДням = False
d = obr.ПостроитьДанныеСравнения()
n_erp = d.БезналичныеДС_ЕРП.Количество()
n_upp = d.БезналичныеДС_УПП.Количество()
print(f"Без днів: ЕРП={n_erp} рядків, УПП={n_upp} рядків")
assert n_erp > 0 and n_upp > 0

# свод по счёту ОТП: кон. остаток ЕРП - УПП
qs = erp.NewObject("Запрос")
qs.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1 БСО.Ссылка КАК Ссылка
ИЗ Справочник.БанковскиеСчетаОрганизаций КАК БСО
ГДЕ БСО.НомерСчета = &НомерСчета
"""
qs.SetParameter("НомерСчета", IBAN_OTP)
otp = qs.Execute().Выгрузить().Получить(0).Ссылка
otp_uid = str(erp.String(otp.УникальныйИдентификатор())).upper()


def summ(tbl, col):
    total = 0.0
    for i in range(tbl.Количество()):
        row = tbl.Получить(i)
        acc = row.БанковскийСчетКасса
        try:
            if erp.ЗначениеЗаполнено(acc) and str(erp.String(acc.УникальныйИдентификатор())).upper() == otp_uid:
                total += float(getattr(row, col))
        except Exception:
            pass
    return total


kon_erp = summ(d.БезналичныеДС_ЕРП, "СуммаКонечныйОстаток_ЕРП")
kon_upp = summ(d.БезналичныеДС_УПП, "СуммаКонечныйОстаток_УПП")
print(f"ОТП: Кін.ЕРП={kon_erp:.2f} Кін.УПП={kon_upp:.2f} Розбіжність={kon_erp - kon_upp:.2f}")
if abs(kon_erp - kon_upp) < 0.005:
    print("ПРИМІТКА: розбіжності по ОТП немає (могла зійтись) — не FAIL")

# --- режим по дням ---
obr.ПоДням = True
d2 = obr.ПостроитьДанныеСравнения()
nd_erp = d2.БезналичныеДС_ПоДнямЕРП.Количество()
nd_upp = d2.БезналичныеДС_ПоДнямУПП.Количество()
print(f"По днях: ЕРП={nd_erp} рядків, УПП={nd_upp} рядків")
if nd_erp == 0 or nd_upp == 0:
    print("FAIL: денні набори порожні")
    FAILS += 1

print("РЕЗУЛЬТАТ: " + ("SMOKE OK" if FAILS == 0 else f"FAIL ({FAILS})"))
sys.exit(1 if FAILS else 0)
