# -*- coding: utf-8 -*-
# Rule #-1 discovery: новый запрос для ТЧ НачисленияБухЗасчетФондов.
# Проверяем фильтр ПоказателиВидовОплатыТруда = ОтчетПоТруду_Соцстрах (новый Q4)
# и партицию: Q3(<>) + Q4(=) == Total(без индикатора)?  + сравнение со старым Код="00021".
# Запуск: C:\Python313\python.exe <этот файл>
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
zup = v8.Connect('Srvr="localhost";Ref="zup";Usr="cfo";Pwd="2442"')
print("Connected ERP + zup_1")

qOrg = erp.NewObject("Запрос")
qOrg.Текст = ('ВЫБРАТЬ Организации.КодПоЕДРПОУ КАК КодПоЕДРПОУ '
             'ИЗ Справочник.Организации КАК Организации '
             'ГДЕ Организации.А_ВБалансе = ИСТИНА И Организации.КодПоЕДРПОУ <> ""')
selOrg = qOrg.Выполнить().Выбрать()
edrpou = []
while selOrg.Следующий():
    edrpou.append(selOrg.КодПоЕДРПОУ)
arr = zup.NewObject("Массив")
for e in edrpou:
    arr.Добавить(e)
print("ЕДРПОУ orgs (А_ВБалансе):", len(edrpou))

BASE = """ВЫБРАТЬ
	Нач.Сотрудник.Физлицо.КодПоДРФО КАК СотрудникИНН,
	Нач.Сотрудник.Физлицо.Наименование КАК СотрудникНаименование,
	Нач.ВидРасчета.Наименование КАК ВидРасчета,
	Нач.ВидРасчета.Код КАК КодВидРасчета,
	СУММА(Нач.Результат) КАК Сумма,
	ВложенныйЗапрос.КодПоЕДРПОУ КАК КодПоЕДРПОУ
ИЗ
	РегистрРасчета.ОсновныеНачисленияРаботниковОрганизаций КАК Нач
		ВНУТРЕННЕЕ СОЕДИНЕНИЕ (ВЫБРАТЬ
			КодыОрганизацииСрезПоследних.Организация КАК Организация,
			КодыОрганизацииСрезПоследних.КодПоЕДРПОУ КАК КодПоЕДРПОУ
		ИЗ
			РегистрСведений.КодыОрганизации.СрезПоследних КАК КодыОрганизацииСрезПоследних) КАК ВложенныйЗапрос
		ПО Нач.Организация = ВложенныйЗапрос.Организация
ГДЕ
	Нач.ПериодРегистрации МЕЖДУ &НачалоПериода И &КонецПериода
	%FILTER%
	И Нач.Результат <> 0
	И ВложенныйЗапрос.КодПоЕДРПОУ В(&масКодПоЕДРПОУ)
СГРУППИРОВАТЬ ПО
	Нач.Сотрудник.Физлицо.КодПоДРФО,
	Нач.Сотрудник.Физлицо.Наименование,
	Нач.ВидРасчета.Наименование,
	Нач.ВидРасчета.Код,
	ВложенныйЗапрос.КодПоЕДРПОУ"""

INDIC = 'Нач.ВидРасчета.ПоказателиВидовОплатыТруда'
SOC = 'ЗНАЧЕНИЕ(Справочник.СтатьиНалоговыхДеклараций.ОтчетПоТруду_Соцстрах)'
FILTERS = {
    "Q4new(=Соцстрах)":   f'И {INDIC} = {SOC}',
    "Q3live(<>Соцстрах)": f'И {INDIC} <> {SOC}',
    "Total(no indic)":    '',
    "OldQ4(=00021)":      'И Нач.ВидРасчета.Код = "00021"',
}

def run(ftext, nm, km):
    q = zup.NewObject("Запрос")
    q.Текст = BASE.replace("%FILTER%", ftext)
    q.УстановитьПараметр("НачалоПериода", nm)
    q.УстановитьПараметр("КонецПериода", km)
    q.УстановитьПараметр("масКодПоЕДРПОУ", arr)
    try:
        r = q.Выполнить().Выгрузить()
        n = r.Количество()
        s = 0.0; vids = {}; rows = []
        for i in range(n):
            row = r.Получить(i)
            s += float(row.Сумма)
            vids[str(row.КодВидРасчета)] = str(row.ВидРасчета)
            rows.append((str(row.СотрудникИНН), str(row.СотрудникНаименование), str(row.КодВидРасчета), float(row.Сумма)))
        return n, s, vids, rows
    except Exception as e:
        msg = e.excepinfo[2] if (hasattr(e, 'excepinfo') and e.excepinfo) else str(e)
        return None, msg, None, None

def month(y, m):
    nm = datetime.datetime(y, m, 1, 0, 0, 0)
    km = (datetime.datetime(y, 12, 31, 23, 59, 59) if m == 12
          else datetime.datetime(y, m + 1, 1) - datetime.timedelta(seconds=1))
    print(f"\n===== {y}-{m:02d} =====")
    res = {}
    detail_q4 = None
    for label, ftext in FILTERS.items():
        n, s, vids, rows = run(ftext, nm, km)
        if n is None:
            print(f"  {label:22} FAIL: {s}")
        else:
            print(f"  {label:22} rows={n:4} Σ={s:,.2f}")
            res[label] = (n, s, vids)
            if label == "Q4new(=Соцстрах)":
                detail_q4 = (vids, rows)
    if all(k in res for k in ["Q4new(=Соцстрах)", "Q3live(<>Соцстрах)", "Total(no indic)"]):
        q4 = res["Q4new(=Соцстрах)"][1]; q3 = res["Q3live(<>Соцстрах)"][1]; tot = res["Total(no indic)"][1]
        print(f"  PARTITION: Q3(<>)+Q4(=) = {q3+q4:,.2f}  Total = {tot:,.2f}  GAP = {tot-(q3+q4):,.2f}")
    if detail_q4:
        vids, rows = detail_q4
        print(f"  Q4 виды расчёта ({len(vids)}): {vids}")
        skr = [r for r in rows if 'крипник' in r[1]]
        if skr:
            print(f"  SANITY Скрипник: {skr}")
        rows.sort(key=lambda r: -r[3])
        print("  Q4 top-6 строк:")
        for r in rows[:6]:
            print(f"     ИНН={r[0]} | {r[1]} | код={r[2]} | {r[3]:,.2f}")

for (y, m) in [(2025, 12), (2026, 1)]:
    month(y, m)

zup = None; erp = None
print("\nDone.")
