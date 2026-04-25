# -*- coding: utf-8 -*-
# Інтеграційний тест процедури ЗаполнитьНачисленияУпризЗуп2()
# у документі А_ОтражениеЗПпоКазне — заповнює НачисленияУпр + УдержанияУпр з zup_2.
#
# ПЕРЕДУМОВА: оновлений ObjectModule.bsl завантажений у конфігурацію (F7).
import win32com.client, sys, pywintypes, datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String


def com_date(y, m, d, hh=0, mm=0, ss=0):
    return pywintypes.Time(datetime.datetime(y, m, d, hh, mm, ss))


def parse_num(s):
    return float(S(s).replace('\xa0', '').replace(' ', '').replace(',', '.') or "0")


# === Знайти документ за грудень 2025 ===
q = erp.NewObject("Запрос")
q.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка "
    "ИЗ Документ.А_ОтражениеЗПпоКазне КАК Д "
    "ГДЕ Д.Дата МЕЖДУ &Н И &К И НЕ Д.ПометкаУдаления "
    "УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ"
)
q.SetParameter("Н", com_date(2025, 12, 1))
q.SetParameter("К", com_date(2025, 12, 31, 23, 59, 59))
sel = q.Execute().Выбрать()
if not sel.Следующий():
    print("ПОМИЛКА: не знайдено А_ОтражениеЗПпоКазне за грудень 2025.")
    sys.exit(1)

obj = sel.Ссылка.ПолучитьОбъект()
print(f"Документ: {S(obj.Ссылка)}")
print(f"Дата: {S(obj.Дата)}, Організация: {S(obj.Организация)}")

# === Стан ДО ===
nch_before = obj.НачисленияУпр.Количество()
ud_before  = obj.УдержанияУпр.Количество()
print(f"\nДО виклику:")
print(f"  НачисленияУпр: {nch_before}")
print(f"  УдержанияУпр:  {ud_before}")

# === Виклик ===
print("\n>>> Виклик obj.ЗаполнитьНачисленияУпризЗуп2()...")
try:
    obj.ЗаполнитьНачисленияУпризЗуп2()
    print(">>> OK")
except Exception as e:
    print(f">>> ПОМИЛКА: {e}")
    sys.exit(1)

# === Стан ПІСЛЯ ===
nch_after = obj.НачисленияУпр.Количество()
ud_after  = obj.УдержанияУпр.Количество()
print(f"\nПІСЛЯ виклику:")
print(f"  НачисленияУпр: {nch_after}")
print(f"  УдержанияУпр:  {ud_after}")

# === Контрольні суми ===
total_nch = 0.0
distinct_nch_types = set()
for i in range(nch_after):
    r = obj.НачисленияУпр.Получить(i)
    total_nch += parse_num(r.Сумма)
    distinct_nch_types.add(S(r.ВидРасчета))

total_ud = 0.0
distinct_ud_types = set()
for i in range(ud_after):
    r = obj.УдержанияУпр.Получить(i)
    total_ud += parse_num(r.Сумма)
    distinct_ud_types.add(S(r.ВидРасчета))

print(f"\nКонтрольні суми:")
print(f"  Сума НачисленияУпр: {total_nch:,.2f}")
print(f"  Сума УдержанияУпр:  {total_ud:,.2f}")
print(f"\n  Види розрахунку Начислений ({len(distinct_nch_types)}):")
for t in sorted(distinct_nch_types):
    print(f"    - {t}")
print(f"\n  Види розрахунку Удержаний ({len(distinct_ud_types)}):")
for t in sorted(distinct_ud_types):
    print(f"    - {t}")

# === Перевірки структури (перші 3 рядки кожної ТЧ) ===
print(f"\n=== Перші 3 рядки НачисленияУпр ===")
for i in range(min(3, nch_after)):
    r = obj.НачисленияУпр.Получить(i)
    print(f"  [{i+1}] {S(r.ФизЛицо):40s} | {S(r.ВидРасчета)[:30]:30s} | код={S(r.КодВидРасчета)} | сума={parse_num(r.Сумма):>10.2f}")

print(f"\n=== Перші 3 рядки УдержанияУпр ===")
for i in range(min(3, ud_after)):
    r = obj.УдержанияУпр.Получить(i)
    print(f"  [{i+1}] {S(r.ФизЛицо):40s} | {S(r.ВидРасчета)[:30]:30s} | код={S(r.КодВидРасчета)} | сума={parse_num(r.Сумма):>10.2f}")

# === Перевірки ===
print("\n=== ПЕРЕВІРКИ ===")
checks = [
    (nch_after > 0, "НачисленияУпр не пуста"),
    (ud_after > 0, "УдержанияУпр не пуста"),
]

if nch_after > 0:
    all_fl = all(erp.ЗначениеЗаполнено(obj.НачисленияУпр.Получить(i).ФизЛицо) for i in range(nch_after))
    all_sum = all(parse_num(obj.НачисленияУпр.Получить(i).Сумма) != 0 for i in range(nch_after))
    checks.append((all_fl, "Усі рядки НачисленияУпр мають ФизЛицо"))
    checks.append((all_sum, "Усі рядки НачисленияУпр мають Сумма <> 0"))

if ud_after > 0:
    all_fl_ud = all(erp.ЗначениеЗаполнено(obj.УдержанияУпр.Получить(i).ФизЛицо) for i in range(ud_after))
    all_sum_ud = all(parse_num(obj.УдержанияУпр.Получить(i).Сумма) != 0 for i in range(ud_after))
    checks.append((all_fl_ud, "Усі рядки УдержанияУпр мають ФизЛицо"))
    checks.append((all_sum_ud, "Усі рядки УдержанияУпр мають Сумма <> 0"))

all_ok = True
for ok, desc in checks:
    mark = "OK" if ok else "FAIL"
    print(f"  [{mark}] {desc}")
    if not ok:
        all_ok = False

print()
if all_ok:
    print("==> ВСІ ПЕРЕВІРКИ ПРОЙДЕНО")
    sys.exit(0)
else:
    print("==> Є ПОМИЛКИ")
    sys.exit(1)
