# -*- coding: utf-8 -*-
# Перевірка фікса фін-агента: Сумма ЕРП заповнюється з А_ПриходДенегОтФинАгента.Сумма.
# Період січень 2026, каса "2 Касса Строительство" (де у скріншоті були порожні Сума ЕРП).
import sys, datetime
import win32com.client, pywintypes
sys.stdout.reconfigure(encoding="utf-8")

EPF = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\trusting-noyce-fbeddb\_Rarzrabotki\Обработки\СинхронизироватьДеньгиКасса.epf"
CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

def err(e):
    if hasattr(e, "excepinfo") and e.excepinfo: return str(e.excepinfo[2])
    return str(e)

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)
proc = erp.ВнешниеОбработки.Создать(EPF, False)
proc.НачалоПериода = pywintypes.Time(datetime.datetime(2026, 1, 1, 12, 0, 0))
proc.ОкончаниеПериода = pywintypes.Time(datetime.datetime(2026, 1, 31, 12, 0, 0))
proc.ФильтрКасса = erp.Справочники.Кассы.НайтиПоНаименованию("2 Касса Строительство")
print("Каса фільтр:", erp.String(proc.ФильтрКасса))

print("СравнитьОстатки:", erp.String(proc.СравнитьОстатки()))
тр = proc.ТаблицаРасхождений
print("Рядків розбіжностей:", тр.Количество())
if тр.Количество() == 0:
    print("Каса звірена за залишками — для аналізу документів виберемо її примусово через індекс 0 неможливо.")
    sys.exit(0)

print("АнализироватьДокументы(0):", erp.String(proc.АнализироватьДокументы(0)))
тд = proc.ТаблицаДокументов
print("Рядків документів:", тд.Количество())
fin = 0
empty_erp = 0
for i in range(тд.Количество()):
    s = тд.Получить(i)
    док = erp.String(s.ДокументЕРП)
    if "фин агент" in док.lower() or "фінагент" in erp.String(s.Статус).lower():
        fin += 1
        if float(s.СуммаЕРП) == 0:
            empty_erp += 1
        print(f"  ФIНАГЕНТ: ДокЕРП={док[:42]:42} | СумЕРП={float(s.СуммаЕРП):>12.2f} | СумКазна={float(s.СуммаКазна):>12.2f}"
              f" | Дія='{erp.String(s.Действие)}' | {erp.String(s.Статус)[:40]}")
print(f"\nПІДСУМОК: фін-агент рядків={fin}, з них порожня Сума ЕРП={empty_erp} (очікується 0)")
print("ГОТОВО.")
