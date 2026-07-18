# -*- coding: utf-8 -*-
# Інтеграційний тест процедури ЗагрузитьНалогиИУдержанияИзЗарплатыБух()
# у документі А_ОтражениеЗПпоКазне.
#
# ПЕРЕДУМОВА: оновлений ObjectModule.bsl завантажений у конфігурацію
# (Конфігуратор → Загрузить конфигурацию из файлов → F7).
import win32com.client, sys, pywintypes, datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String


def com_date(y, m, d, hh=0, mm=0, ss=0):
    return pywintypes.Time(datetime.datetime(y, m, d, hh, mm, ss))


# === Крок 1: знайти документ А_ОтражениеЗПпоКазне за грудень 2025 ===
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
    print("ПОМИЛКА: не знайдено А_ОтражениеЗПпоКазне за грудень 2025. Створіть документ і запустіть знову.")
    sys.exit(1)

obj = sel.Ссылка.ПолучитьОбъект()
print(f"Документ: {S(obj.Ссылка)}")
print(f"Дата: {S(obj.Дата)}, Організация: {S(obj.Организация)}")

# === Крок 2: стан ДО ===
nb_before = obj.НалогиБухгалтерия.Количество()
ub_before = obj.УдержанияБухгалтерия.Количество()
print(f"\nДО виклику:")
print(f"  НалогиБухгалтерия:   {nb_before}")
print(f"  УдержанияБухгалтерия: {ub_before}")

# === Крок 3: виклик процедури ===
print("\n>>> Виклик obj.ЗагрузитьНалогиИУдержанияИзЗарплатыБух()...")
try:
    obj.ЗагрузитьНалогиИУдержанияИзЗарплатыБух()
    print(">>> OK")
except Exception as e:
    print(f">>> ПОМИЛКА: {e}")
    sys.exit(1)

# === Крок 4: стан ПІСЛЯ ===
nb_after = obj.НалогиБухгалтерия.Количество()
ub_after = obj.УдержанияБухгалтерия.Количество()
print(f"\nПІСЛЯ виклику:")
print(f"  НалогиБухгалтерия:   {nb_after}")
print(f"  УдержанияБухгалтерия: {ub_after}")

# === Крок 5: деталі УдержанияБухгалтерия ===
print("\n=== УдержанияБухгалтерия ===")
total_ud = 0.0
ud_rows = []
for i in range(ub_after):
    r = obj.УдержанияБухгалтерия.Получить(i)
    sum_v = float(S(r.Сумма).replace('\xa0', '').replace(' ', '').replace(',', '.') or "0")
    total_ud += sum_v
    ud_rows.append({
        "ФизЛицо":     S(r.ФизическоеЛицо),
        "ВидОпер":     S(r.ВидОперации),
        "Статья":      S(r.СтатьяАктивовПассивов),
        "Сумма":       sum_v,
        "Счет":        S(r.СчетУчета),
        "Субк1":       S(r.Субконто1),
        "Группа":      S(r.ГруппаУчетаНачислений),
    })
    print(f"\n[{i+1}] {ud_rows[-1]['ФизЛицо']}")
    print(f"    ВидОперации     : {ud_rows[-1]['ВидОпер']}")
    print(f"    Статья          : {ud_rows[-1]['Статья']}")
    print(f"    СчетУчета       : {ud_rows[-1]['Счет']}")
    print(f"    Субконто1       : {ud_rows[-1]['Субк1']}")
    print(f"    Группа учета    : {ud_rows[-1]['Группа']}")
    print(f"    Сумма           : {sum_v:>10.2f}")

print(f"\nРазом удержаний: {total_ud:.2f}")

# === Перевірки ===
print("\n=== ПЕРЕВІРКИ ===")

checks = [
    (nb_after > 0, "НалогиБухгалтерия не пуста"),
    (ub_after > 0, "УдержанияБухгалтерия не пуста"),
]

# Кожен рядок удержань має правильну структуру
if ub_after > 0:
    all_have_fl = all(bool(r['ФизЛицо']) for r in ud_rows)
    all_6853 = all(r['Счет'].strip() == "6853" for r in ud_rows)
    all_proche_passives = all("Прочие" in r['Статья'] or "Прочі" in r['Статья'] for r in ud_rows)
    all_have_subk1 = all(bool(r['Субк1']) for r in ud_rows)
    all_group = all(("Зарплата" in r['Группа'] or "661" in r['Группа']) for r in ud_rows)
    checks.append((all_have_fl, "Усі рядки удержань мають ФизическоеЛицо"))
    checks.append((all_6853, "Усі рядки удержань мають СчетУчета=6853"))
    checks.append((all_proche_passives, "Усі рядки удержань мають СтатьяАктивовПассивов='Прочие пассивы'"))
    checks.append((all_have_subk1, "Усі рядки удержань мають Субконто1 (Контрагент)"))
    checks.append((all_group, "Усі рядки удержань мають ГруппаУчетаНачислений='Зарплата'"))

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
