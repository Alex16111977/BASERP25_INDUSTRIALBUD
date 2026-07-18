# -*- coding: utf-8 -*-
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# Статья Налоги
qst = erp.NewObject("Запрос")
qst.Text = """
ВЫБРАТЬ Ст.Ссылка КАК Ссылка, Ст.Наименование КАК Наим
ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов КАК Ст
ГДЕ Ст.Наименование ПОДОБНО "%алог%"
"""
print("=== Статьи АктивовПассивов с 'налог' ===")
ст_налоги = None
for s in qst.Execute().Выгрузить():
    print(f"  '{s.Наим}'")
    if s.Наим == "Налоги":
        ст_налоги = s.Ссылка

# Перечисление ТипыНалогов
print("\n=== Значения Перечисление.ТипыНалогов ===")
for зн in erp.Перечисления.ТипыНалогов:
    print(f"  {erp.XMLСтрока(зн)}")

# Управленческий регистр: остатки по статье Налоги на 31.12.2025, разрез Орг x Аналитика(ТипНалога)
def test_upr(d, ст):
    moment = datetime.datetime(d.year, d.month, d.day, 23, 59, 59)
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ
        Ост.Организация.Наименование КАК Орг,
        Ост.Организация.КодПоЕДРПОУ КАК ЕДРПОУ,
        ПРЕДСТАВЛЕНИЕ(Ост.Аналитика) КАК Налог,
        Ост.СуммаОстаток КАК Сальдо
    ИЗ
        РегистрНакопления.ПрочиеАктивыПассивы.Остатки(&Момент, Статья = &Статья) КАК Ост
    ГДЕ Ост.СуммаОстаток <> 0
    УПОРЯДОЧИТЬ ПО ЕДРПОУ
    """
    q.SetParameter("Момент", moment)
    q.SetParameter("Статья", ст)
    try:
        r = q.Execute().Выгрузить()
        print(f"\n=== ERP УПР ПрочиеАктивыПассивы статья Налоги на {d.strftime('%d.%m.%Y')} — рядкiв={r.Количество()} ===")
        итог = 0
        for s in r:
            print(f"  ЕДРПОУ={str(s.ЕДРПОУ):11} | {str(s.Налог):20} | Сальдо={s.Сальдо:>15,.2f} | {s.Орг[:22]}")
            итог += s.Сальдо
        print(f"  ИТОГО={итог:,.2f}")
    except Exception as e:
        if hasattr(e, 'excepinfo') and e.excepinfo:
            print(f"  FAIL: {e.excepinfo[2]}")
        else:
            print(f"  FAIL: {e}")

if ст_налоги:
    for d in [datetime.date(2025,11,30), datetime.date(2025,12,31)]:
        test_upr(d, ст_налоги)
else:
    print("Статья 'Налоги' не найдена точным именем")
