# -*- coding: utf-8 -*-
"""ПАП «Оплата труда» — какие типы документов пишут каждую «ногу» (empty vs named источник)."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def f(x):
    try: return float(x)
    except: return 0.0

ОТ = erp.NewObject("Запрос")
ОТ.Текст = ("ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов "
            "ГДЕ Наименование = \"Оплата труда\"")
ОТ = ОТ.Выполнить().Выгрузить()[0].С

q = erp.NewObject("Запрос")
q.УстановитьПараметр("ОТ", ОТ)
q.Текст = """
ВЫБРАТЬ
    Т.Регистратор КАК Рег,
    ВЫБОР КОГДА Т.Источник = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.ПустаяСсылка)
        ТОГДА "empty" ИНАЧЕ "named" КОНЕЦ КАК Нога,
    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход) ТОГДА Т.Сумма ИНАЧЕ -Т.Сумма КОНЕЦ) КАК КО
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
ГДЕ Т.Период <= ДАТАВРЕМЯ(2026,1,31,23,59,59) И Т.Статья = &ОТ
СГРУППИРОВАТЬ ПО Т.Регистратор,
    ВЫБОР КОГДА Т.Источник = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.ПустаяСсылка)
        ТОГДА "empty" ИНАЧЕ "named" КОНЕЦ
"""
res = q.Выполнить().Выгрузить()

# агрегируем по (тип документа, нога)
buckets = {}
typecache = {}
for r in res:
    ref = r.Рег
    key = erp.XMLТип(ref).ИмяТипа if False else None  # placeholder
    try:
        tn = ref.Метаданные().Имя
    except Exception:
        tn = "?"
    nm = str(r.Нога)
    k = (tn, nm)
    buckets[k] = buckets.get(k, 0.0) + f(r.КО)

print(f"{'Тип документа-регистратора':<48} | {'нога':<6} | {'Σ КО':>16}")
for (tn, nm) in sorted(buckets, key=lambda x: (x[1], -abs(buckets[x]))):
    print(f"  {tn:<48} | {nm:<6} | {buckets[(tn,nm)]:>16,.2f}")

# итоги по ноге
print("\n--- итоги по ноге ---")
legtot = {}
for (tn, nm), v in buckets.items():
    legtot[nm] = legtot.get(nm, 0.0) + v
for nm, v in legtot.items():
    print(f"  {nm:<6}: {v:>16,.2f}")
print(f"  NET : {sum(legtot.values()):>16,.2f}")
