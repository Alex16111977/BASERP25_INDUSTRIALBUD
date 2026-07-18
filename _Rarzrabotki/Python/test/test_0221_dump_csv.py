# -*- coding: utf-8 -*-
"""Выгрузка полного перечня резидуумов 0221 (Кол=0, Сумма<>0) в CSV. Rule #-1."""
import win32com.client, sys, datetime, csv
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')
def S(x):
    try: return buh.String(x)
    except Exception: return str(x)

счет = buh.ПланыСчетов.Хозрасчетный.НайтиПоКоду("0221").Ссылка
q = buh.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ
    ПРЕДСТАВЛЕНИЕ(Ост.Организация) КАК Организация,
    ПРЕДСТАВЛЕНИЕ(Ост.Субконто1) КАК Контрагент,
    ПРЕДСТАВЛЕНИЕ(Ост.Субконто2) КАК Номенклатура,
    ПРЕДСТАВЛЕНИЕ(Ост.Субконто3) КАК Склад,
    Ост.СуммаОстатокДт КАК Сумма,
    Ост.КоличествоОстатокДт КАК Количество
ИЗ РегистрБухгалтерии.Хозрасчетный.Остатки(&НаДату, Счет = &Счет) КАК Ост
ГДЕ Ост.КоличествоОстатокДт = 0 И Ост.СуммаОстатокДт <> 0
УПОРЯДОЧИТЬ ПО Сумма УБЫВ
"""
q.УстановитьПараметр("Счет", счет)
q.УстановитьПараметр("НаДату", datetime.datetime(2026, 7, 13, 23, 59, 59))
r = q.Выполнить().Выгрузить()

путь = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\prompts\diagnostika_0221_residuumy.csv"
итого = 0.0
with open(путь, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow(["№", "Организация", "Контрагент", "Номенклатура", "Склад", "Сумма_БО", "Количество"])
    for i in range(r.Количество()):
        s = r.Получить(i)
        итого += float(s.Сумма)
        w.writerow([i+1, S(s.Организация), S(s.Контрагент), S(s.Номенклатура), S(s.Склад),
                    f"{float(s.Сумма):.2f}".replace(".", ","), f"{float(s.Количество):.0f}"])
    w.writerow([])
    w.writerow(["ИТОГО", "", "", "", "", f"{итого:.2f}".replace(".", ","), "0"])
print(f"Записано строк: {r.Количество()} | нетто-сумма резидуумов: {итого:.2f}")
print(f"CSV: {путь}")
