# -*- coding: utf-8 -*-
"""Инварианты колонок оплаты в отчёте «Прогноз виконання СС за потребою» (встроенный, BaseERP). Только чтение.

    python test_oplata_v_otchete.py [ГГГГ-ММ-ДД]
"""
import datetime as dt
import sys

import win32com.client

sys.stdout.reconfigure(encoding='utf-8')
erp = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
d = dt.datetime.strptime(sys.argv[1], "%Y-%m-%d") if len(sys.argv) > 1 else dt.datetime.now()
tz = erp.Отчеты.А_ПрогнозВиконанняССЗаПотребою.Создать().ПолучитьДанные(d + dt.timedelta(hours=12))
OLD = ["ПланКол", "ПланГрн", "КупленоКол", "КупленоГрн", "ПослеГрн", "ЗалишокСС", "ПотребаКол", "ПотребаГрн",
       "ЗалишокПотребы", "НеДокупаем"]
ok = True
tot = {k: 0.0 for k in ("ОплаченоНеНадійшло", "Оплачено", "Борг", "КупленоЗак")}
bad = []


def check(c, m):
    global ok
    print(("OK   " if c else "FAIL ") + m)
    ok = ok and bool(c)


for i in range(tz.Количество()):
    r = tz.Получить(i)
    vid = S(r.ВидРядка)
    onn, opl, borg = float(r.ОплаченоНеНадійшло), float(r.Оплачено), float(r.Борг)
    tot["ОплаченоНеНадійшло"] += onn
    if vid == "Рахунок оплачений, не надійшов":
        if any(float(getattr(r, k)) for k in OLD) or onn <= 0:
            bad.append(("строка счёта со старыми ресурсами или ≤0", S(r.НазваЗагальна)))
    elif onn:
        bad.append(("ОплаченоНеНадійшло вне строк счёта", vid))
    if vid.startswith("Закупівля"):
        if opl or borg:
            tot["Оплачено"] += opl
            tot["Борг"] += borg
            tot["КупленоЗак"] += float(r.КупленоГрн)
            if abs(opl + borg - float(r.КупленоГрн)) > 0.011 or opl < -0.01 or borg < -0.011:
                bad.append(("Оплачено+Борг≠Куплено или <0", S(r.Документ)))
    elif opl or borg:
        bad.append(("Оплачено/Борг вне закупок", vid))
z = erp.NewObject("Запрос")
z.Text = "ВЫБРАТЬ ЕСТЬNULL(СУММА(Р.Сумма), 0) ИЗ РегистрСведений.А_ТоварыСчетовПоставщиков КАК Р"
reg = float(z.Execute().Выгрузить().Получить(0).Получить(0))
print(f"строк отчёта {tz.Количество()}; Оплачено, не надійшло {tot['ОплаченоНеНадійшло']:,.2f}; "
      f"закупки с заказом {tot['КупленоЗак']:,.2f} = оплачено {tot['Оплачено']:,.2f} + борг {tot['Борг']:,.2f}".replace(",", " "))
check(not bad, f"нарушений по строкам: {len(bad)} {bad[:3]}")
check(abs(tot["Оплачено"] + tot["Борг"] - tot["КупленоЗак"]) <= 1, "Σ Оплачено + Σ Борг = Σ Куплено (закупки с заказом)")
check(0 <= tot["ОплаченоНеНадійшло"] <= reg + 0.01, f"0 ≤ Оплачено,не надійшло ≤ Σ строк счетов ({reg:,.2f})".replace(",", " "))
sys.exit(0 if ok else 1)
