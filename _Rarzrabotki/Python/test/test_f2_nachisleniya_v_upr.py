# -*- coding: utf-8 -*-
"""
Приёмочный тест: РаспределениеФ2 (А_Начисление=Истина) → А_ОтражениеЗПпоКазне.НачисленияУпр.

Заполняет реальный документ ОЗФУ за дек.2025 В ПАМЯТИ (БЕЗ записи/проведения!)
и проверяет, что:
  1. zup_2-строки сохранены (у Постернака есть строка-«Оклад», НЕ только премия);
  2. появились Ф2-строки с кодом «ПРЕМ» (премии Форма2);
  3. у Постернака есть «Премія (Зарплата)» / ПРЕМ / 100 000;
  4. Σ премий ПРЕМ за период = 2 209 500 (эталон №000000034).

НЕ записывает и НЕ перепроводит документ (LESSONS §19).
"""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN_ERP)

# 1. Найти документ ОЗФУ за дек.2025
q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    А.Ссылка КАК Ссылка, А.Номер КАК Номер, А.Дата КАК Дата, А.Проведен КАК Проведен
ИЗ Документ.А_ОтражениеЗПпоКазне КАК А
ГДЕ А.Дата >= ДАТАВРЕМЯ(2025,12,1,0,0,0) И А.Дата <= ДАТАВРЕМЯ(2025,12,31,23,59,59)
УПОРЯДОЧИТЬ ПО А.Дата
"""
sel = q.Execute().Выбрать()
docs = []
while sel.Следующий():
    docs.append((str(sel.Номер), str(sel.Дата), bool(sel.Проведен), sel.Ссылка))

print("Документи ОЗФУ дек.2025:", [(d[0], d[1], "пров" if d[2] else "не пров") for d in docs])
if not docs:
    print("FAIL: немає документа ОЗФУ за дек.2025")
    sys.exit(1)

nomer, data, _, ref = docs[0]
obj = ref.ПолучитьОбъект()
print(f"\nЗаповнюємо №{nomer} от {data} В ПАМ'ЯТІ (без запису)...")

obj.ЗаполнитьНачисленияУпризЗуп2()

# 2. Прочитать НачисленияУпр
total = 0.0
prem_rows = []
postern = []
n = 0
for row in obj.НачисленияУпр:
    n += 1
    summ = float(row.Сумма)
    total += summ
    kod = str(row.КодВидРасчета).strip()
    fio = str(row.ФИО)
    if kod == "ПРЕМ":
        prem_rows.append((fio, summ))
    if "Постернак" in fio:
        postern.append((str(row.ВидРасчета).strip(), kod, summ))

sum_prem = sum(s for _, s in prem_rows)
print(f"\nНачисленияУпр: {n} рядків, Σ={total:,.2f}")
print(f"  ПРЕМ (Форма2/премії): {len(prem_rows)} рядків, Σ={sum_prem:,.2f}")
print(f"  Постернак рядки: {postern}")

# 3. Проверки
ok = True
if abs(sum_prem - 2209500) > 0.01:
    print(f"FAIL: Σ ПРЕМ={sum_prem}, очікувалось 2 209 500"); ok = False
if not any(k == "ПРЕМ" and abs(s - 100000) < 0.01 for _, k, s in postern):
    print("FAIL: немає рядка Постернак ПРЕМ 100 000"); ok = False
if not any(k != "ПРЕМ" for _, k, s in postern):
    print("FAIL: немає zup_2-рядка Постернака (Оклад) — zup_2-рядки затёрто Ф2-кроком!"); ok = False
if len(prem_rows) != 44:
    print(f"WARN: ПРЕМ рядків {len(prem_rows)}, очікувалось 44 (за №000000034)")

print("\nOK: всі перевірки пройдені" if ok else "\nFAIL: є помилки (див. вище)")
sys.exit(0 if ok else 1)
