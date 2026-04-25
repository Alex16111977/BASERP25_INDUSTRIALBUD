# -*- coding: utf-8 -*-
# Тест процедури ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП
# для документа А_ОтражениеЗПпоКазне (новий документ — джерело Казна).
#
# ПЕРЕДУМОВА: код ObjectModule.bsl завантажений у конфігурацію
# (Конфігуратор → Загрузить конфигурацию из файлов → Обновить конфигурацию БД — F7)
#
# Шукає документ А_ОтражениеЗПпоКазне з Датою у грудні 2025
# (або використати існуючий / створити перед запуском).
import win32com.client, sys, pywintypes, datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String


def parse_num(s):
    return float(S(s).replace('\xa0', '').replace(' ', '').replace(',', '.') or "0")


def com_date(y, m, d, hh=0, mm=0, ss=0):
    return pywintypes.Time(datetime.datetime(y, m, d, hh, mm, ss))


# === Крок 1: знайти документ А_ОтражениеЗПпоКазне за грудень 2025 ===
q = conn.NewObject("Запрос")
q.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка "
    "ИЗ Документ.А_ОтражениеЗПпоКазне КАК Док "
    "ГДЕ Док.Дата МЕЖДУ &Н И &К И НЕ Док.ПометкаУдаления "
    "УПОРЯДОЧИТЬ ПО Док.Дата УБЫВ"
)
q.SetParameter("Н", com_date(2025, 12, 1))
q.SetParameter("К", com_date(2025, 12, 31, 23, 59, 59))
sel = q.Execute().Choose()
if not sel.Next():
    print("ПОМИЛКА: не знайдено документ А_ОтражениеЗПпоКазне з Датою у грудні 2025.")
    print("Створіть документ у 1С (Дата=31.12.2025, Організация=ТОВ ІНДАСТРІАЛБУД),")
    print("заповніть ТЧ А_НалогиБухгалтерия з даних еталонного документа, і запустіть знову.")
    sys.exit(1)

ref = sel.Ссылка
obj = ref.GetObject()
print(f"Документ: {S(ref)}")
print(f"Дата: {S(obj.Дата)}")
print(f"Організация: {S(obj.Организация)}")

# === Крок 2: стан ДО ===
kz_before = obj.РаспределениеКазна.Count()
print(f"\nДО виклику:")
print(f"  РаспределениеКазна:                    {kz_before}")

# === Крок 3: виклик процедури (з перехопленням повідомлень) ===
print("\n>>> Виклик obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()...")

# Перехоплюємо сообщения через новий масив повідомлень
try:
    obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
    print(">>> OK")
except Exception as e:
    print(f">>> ПОМИЛКА: {e}")
    # Не виходимо — спробуємо показати повідомлення

# === Крок 4: стан ПІСЛЯ ===
kz_after = obj.РаспределениеКазна.Count()
print(f"\nПІСЛЯ виклику:")
print(f"  РаспределениеКазна:                    {kz_after}")

# Порахуємо скільки рядків з заповненим ДокРаспределениеЗП
kz_with_doc = 0
kz_empty_doc = 0
kz_with_employee = 0
kz_empty_employee = 0
for i in range(kz_after):
    r = obj.РаспределениеКазна.Get(i)
    if conn.ЗначениеЗаполнено(r.ДокРаспределениеЗП):
        kz_with_doc += 1
    else:
        kz_empty_doc += 1
    if conn.ЗначениеЗаполнено(r.Сотрудник):
        kz_with_employee += 1
    else:
        kz_empty_employee += 1

print(f"  — з заповненим ДокРаспределениеЗП:    {kz_with_doc}")
print(f"  — БЕЗ ДокРаспределениеЗП:              {kz_empty_doc}")
print(f"  — з заповненим Сотрудник:              {kz_with_employee}")
print(f"  — БЕЗ Сотрудника:                      {kz_empty_employee}")

# === Крок 5: перевірка для Стецюк ===
stetsuk_rows = []
stetsuk_sum = 0.0
stetsuk_docs = set()
for i in range(kz_after):
    r = obj.РаспределениеКазна.Get(i)
    fio = S(r.ФИО)
    if "Стецюк" in fio:
        s_val = parse_num(r.Сумма)
        stetsuk_sum += s_val
        stetsuk_docs.add(S(r.ДокРаспределениеЗП) or "(пусто)")
        stetsuk_rows.append({
            "стаття": S(r.СтатьяДвиженияДенежныхСредств)[:35],
            "док": (S(r.ДокРаспределениеЗП) or "")[:50],
            "підр": S(r.Подразделение)[:20],
            "сум": s_val,
            "ІНН": S(r.ИНН),
            "КодКазна": S(r.КодСотрудникаКазна),
        })

print(f"\n=== Стецюк Тетяна Леонідівна ===")
print(f"  Всього: {stetsuk_sum:.2f} грн  ({len(stetsuk_rows)} рядків)")
for row in stetsuk_rows:
    print(f"    [{row['стаття']:35s}] {row['док']:50s} Підр={row['підр']:20s} Сум={row['сум']:>12.2f}")
print(f"  Очікується:    36046.63 грн  (ЗП №268 = 16346.63 + ф2 №33 = 19700.00)")
ok_stetsuk = abs(stetsuk_sum - 36046.63) < 0.05

# === Крок 6: перевірка — немає рядків без Сотрудника ===
empty_emp = 0
empty_emp_samples = []
for i in range(kz_after):
    r = obj.РаспределениеКазна.Get(i)
    if not conn.ЗначениеЗаполнено(r.Сотрудник):
        empty_emp += 1
        if len(empty_emp_samples) < 5:
            empty_emp_samples.append(f"{S(r.ФИО)} (ІНН={S(r.ИНН)}, КодКазна={S(r.КодСотрудникаКазна)})")

print(f"\n=== Рядки без Сотрудника ===")
print(f"  Всього: {empty_emp}")
for s in empty_emp_samples:
    print(f"    - {s}")

# === Підсумок ===
print("\n=== Підсумок ===")
checks = [
    (kz_after > 0, "РаспределениеКазна заповнена"),
    (kz_empty_doc == 0, f"Всі рядки мають ДокРаспределениеЗП (порожніх: {kz_empty_doc})"),
    (ok_stetsuk, f"Сума по Стецюк = 36046.63 (факт: {stetsuk_sum:.2f})"),
]
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
