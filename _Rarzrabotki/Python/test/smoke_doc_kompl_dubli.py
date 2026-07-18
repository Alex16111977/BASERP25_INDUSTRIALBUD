# -*- coding: utf-8 -*-
"""Smoke 4: РасчетКомплектаций — контроль дублей (та же спецификация/день/пересечение складов).
Второй документ get-or-create по маркеру; запись НЕ блокируется (предупреждение)."""
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MARKER1 = "SMOKE_DOC_KOMPL_v1"
MARKER2 = "SMOKE_DOC_KOMPL_v1_ДУБЛЬ"

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')


def find_by_comment(marker):
    q = buh.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Ссылка ИЗ Документ.РасчетКомплектаций КАК Д "
              "ГДЕ ВЫРАЗИТЬ(Д.Комментарий КАК Строка(200)) = &М И НЕ Д.ПометкаУдаления")
    q.SetParameter("М", marker)
    r = q.Execute().Выгрузить()
    return r.Получить(0).Ссылка if r.Количество() > 0 else None


ref1 = find_by_comment(MARKER1)
assert ref1 is not None, "сначала smoke_doc_kompl_analiz.py"
doc1 = ref1.ПолучитьОбъект()

ref2 = find_by_comment(MARKER2)
if ref2 is not None:
    doc2 = ref2.ПолучитьОбъект()
else:
    doc2 = buh.Документы.РасчетКомплектаций.СоздатьДокумент()
    doc2.Дата = datetime.datetime.now().replace(microsecond=0)
    doc2.Заполнить(None)
    doc2.Комментарий = MARKER2

# те же параметры, что у первого
doc2.Спецификация = doc1.Спецификация
doc2.Период = doc1.Период
doc2.СкладыОстатков.Очистить()
for i in range(doc1.СкладыОстатков.Количество()):
    doc2.СкладыОстатков.Добавить().Склад = doc1.СкладыОстатков.Получить(i).Склад

dupes = doc2.НайтиДубликатыРасчета()
assert dupes.Количество() >= 1, "дубль не найден"
print(f"Найдено дублей: {dupes.Количество()} (первый: {buh.String(dupes.Получить(0))})")

# запись не блокируется (ПередЗаписью — только Сообщить)
doc2.Записать()
print(f"Второй документ записан: №{doc2.Номер} (предупреждение, не отказ)")

# у самого первого документа дубль теперь второй
doc1b = ref1.ПолучитьОбъект()
d2 = doc1b.НайтиДубликатыРасчета()
assert d2.Количество() >= 1
print("Обратная проверка: первый документ видит второй как дубль")
print("DUBLI PASS")
