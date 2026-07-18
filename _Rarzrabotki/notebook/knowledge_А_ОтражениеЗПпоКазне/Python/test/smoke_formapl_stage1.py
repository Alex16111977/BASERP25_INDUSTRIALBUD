# -*- coding: utf-8 -*-
# Smoke Stage 1: РаспределениеКазна.ФормаPL + Организация для Ф2. В память (без записи).
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
print("Connected ERP")

q = erp.NewObject("Запрос")
q.Текст = ("ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК Ссылка, Номер КАК Номер, Дата КАК Дата "
           "ИЗ Документ.А_ОтражениеЗПпоКазне "
           "ГДЕ Дата МЕЖДУ ДАТАВРЕМЯ(2026,4,1,0,0,0) И ДАТАВРЕМЯ(2026,4,30,23,59,59) УПОРЯДОЧИТЬ ПО Дата")
sel = q.Выполнить().Выбрать()
if not sel.Следующий():
    print("FAIL: нет документа"); sys.exit(1)
print(f"Документ: №{sel.Номер} от {sel.Дата}")

obj = sel.Ссылка.ПолучитьОбъект()
obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()

def enum_name(v):
    if not erp.ЗначениеЗаполнено(v):
        return ""
    try:
        return erp.XMLСтрока(v)
    except Exception:
        return "?"

def doc_type(d):
    if not erp.ЗначениеЗаполнено(d):
        return "<нет>"
    try:
        return erp.XMLТип(d).ИмяТипа
    except Exception:
        try:
            return d.Метаданные().Имя
        except Exception:
            return "?"

rk = obj.РаспределениеКазна
n = rk.Количество()
from collections import Counter
cФП = Counter()
ф2_упр = ф2_неупр = 0
рассинхрон = 0
for i in range(n):
    r = rk.Получить(i)
    фп = enum_name(r.ФормаPL)
    cФП[фп] += 1
    тд = doc_type(r.ДокРаспределениеЗП)
    # соответствие маркера и типа документа
    if erp.ЗначениеЗаполнено(r.ДокРаспределениеЗП):
        ожид = "Форма2" if "РаспределениеФ2" in тд else "Форма1"
        if фп != ожид:
            рассинхрон += 1
    if фп == "Форма2":
        edrpou = str(r.Организация.КодПоЕДРПОУ) if erp.ЗначениеЗаполнено(r.Организация) else ""
        if edrpou == "40645273":
            ф2_упр += 1
        else:
            ф2_неупр += 1

print(f"\n=== РаспределениеКазна: {n} строк ===")
for k, v in cФП.items():
    print(f"  ФормаPL '{k or '<пусто>'}' = {v}")
print(f"Ф2 орг=управленч.(40645273) = {ф2_упр} (ожид = все Ф2)")
print(f"Ф2 орг НЕ управленч.        = {ф2_неупр} (ожид 0)")
print(f"рассинхрон маркер↔тип документа = {рассинхрон} (ожид 0)")

# образцы
print("\n--- 4 примера ---")
for i in range(min(4, n)):
    r = rk.Получить(i)
    print(f"  ФормаPL={enum_name(r.ФормаPL)} | типДок={doc_type(r.ДокРаспределениеЗП)} | орг={str(r.Организация)} | Сумма={float(r.Сумма):,.2f}")

erp = None
print("\nDone (не записан).")
