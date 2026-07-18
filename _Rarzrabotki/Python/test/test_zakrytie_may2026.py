# -*- coding: utf-8 -*-
"""Прогон ВыполнитьЗакрытие на период 01.05.2026 - 31.05.2026.

Цель: воспроизвести скриншот пользователя, понять — почему обработка
выдаёт ~100 строк с ошибками, а регламентная операция показывает 1 минус.
"""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\ЗакрытиеОтрицательныхОстатков.epf"
ORG_UUID = "6bee36b2-53f0-11e6-80d3-000c29bbac23"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)

q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ ДАТАВРЕМЯ(2026,5,1) КАК Н, ДАТАВРЕМЯ(2026,5,31) КАК К"
r = q.Execute().Выгрузить().Получить(0)
nach, kon = r.Н, r.К

org = erp.Справочники.Организации.ПолучитьСсылку(
    erp.NewObject("УникальныйИдентификатор", ORG_UUID))

print("=" * 70)
print("ПРОГОН ВыполнитьЗакрытие за Май 2026 (01.05.2026 - 31.05.2026)")
print("=" * 70)

обработка = erp.ВнешниеОбработки.Создать(EPF)
обработка.НачалоПериода = nach
обработка.ОкончаниеПериода = kon
обработка.Организация = org

print("→ .ВыполнитьЗакрытие()")
обработка.ВыполнитьЗакрытие()

тч = обработка.ОтрицательныеОстатки
total = тч.Количество()
print(f"\nТЧ ОтрицательныеОстатки: {total} строк\n")

statuses = {}
err_examples = {}
for i in range(total):
    стр = тч.Получить(i)
    st = str(стр.Статус)
    short = st if len(st) < 60 else st[:57] + "..."
    statuses[short] = statuses.get(short, 0) + 1
    if st.startswith("Ошибка") and st not in err_examples:
        err_examples[st] = (erp.String(стр.Номенклатура), float(стр.Количество))

print("Группировка по статусам:")
for k, v in statuses.items():
    print(f"  [{v}]  {k}")

print("\nПримеры строк с ошибками (полный текст):")
for st, (nom, qty) in list(err_examples.items())[:3]:
    print(f"  Номенклатура: {nom}, qty={qty}")
    print(f"    Статус: {st}")
    print()

# Что фактически вернула остатки
q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К
ИЗ РегистрНакопления.ТоварыОрганизаций.Остатки(&Кон,) КАК Ост
ГДЕ Ост.КоличествоОстаток < 0
"""
q.SetParameter("Кон", erp.NewObject("Запрос").Text)  # dummy
q.Text = """
ВЫБРАТЬ КОЛИЧЕСТВО(РАЗЛИЧНЫЕ Ост.АналитикаУчетаНоменклатуры) КАК К
ИЗ РегистрНакопления.ТоварыОрганизаций.Остатки(ДАТАВРЕМЯ(2026,5,31,23,59,59),) КАК Ост
ГДЕ Ост.КоличествоОстаток < 0
"""
real = q.Execute().Выгрузить().Получить(0).К
print(f"Реально минусов в регистре на 31.05.2026: {real}")
