# -*- coding: utf-8 -*-
"""Pretest4 (Rule #-1): возврат ЕРП-стороны отчёта на РегистрНакопления.ЗарплатаКВыплате
(запрос пользователя 1:1). После правки заполнения НачисленияЗарплаты приход = gross.
Эталон: Постернак (ДРФО 2742610332), декабрь 2025: Начисления = 200 000,00."""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

DRFO = "2742610332"
EXP_NACH = 200000.00

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ
	Ост.ФизическоеЛицо КАК ФизическоеЛицо,
	Ост.ФизическоеЛицо.КодПоДРФО КАК ДРФО,
	СУММА(Ост.СуммаКВыплатеНачальныйОстаток) КАК НачальныйОстаток,
	СУММА(Ост.СуммаКВыплатеПриход) КАК Начисления,
	СУММА(Ост.СуммаКВыплатеРасход) КАК Выплаты,
	СУММА(Ост.СуммаКВыплатеКонечныйОстаток) КАК КонечныйОстаток
ИЗ
	РегистрНакопления.ЗарплатаКВыплате.ОстаткиИОбороты(
		ДАТАВРЕМЯ(2025, 12, 1), ДАТАВРЕМЯ(2025, 12, 31, 23, 59, 59), , , ) КАК Ост
СГРУППИРОВАТЬ ПО
	Ост.ФизическоеЛицо,
	Ост.ФизическоеЛицо.КодПоДРФО"""
try:
    t = q.Execute().Выгрузить()
except Exception as e:
    msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
    print(f"[FAIL] запрос ЗКВ: {msg}")
    sys.exit(1)

print(f"[OK] запрос ЗКВ: строк={t.Количество()}")

found = False
s = [0.0] * 4
for i in range(t.Количество()):
    r = t.Получить(i)
    vals = (float(r.НачальныйОстаток or 0), float(r.Начисления or 0),
            float(r.Выплаты or 0), float(r.КонечныйОстаток or 0))
    for j in range(4):
        s[j] += vals[j]
    if (r.ДРФО or "").strip() == DRFO:
        found = True
        print(f"  Постернак дек2025: НачОст={vals[0]:,.2f} Начисл={vals[1]:,.2f} "
              f"Выпл={vals[2]:,.2f} КонОст={vals[3]:,.2f}")
        ok = abs(vals[1] - EXP_NACH) <= 0.01
        print(f"  CHECK Начисления==200000.00: {'OK' if ok else 'FAIL'}")
        if not ok:
            print("PRETEST4: FAIL")
            sys.exit(1)

print(f"  Σ: НачОст={s[0]:,.2f} Начисл={s[1]:,.2f} Выпл={s[2]:,.2f} КонОст={s[3]:,.2f}")
print("PRETEST4: " + ("PASS" if found else "FAIL (Постернак не найден)"))
sys.exit(0 if found else 1)
