# -*- coding: utf-8 -*-
# Тест PIVOT-агрегації РаспределениеКазна (4 колонки: Сумма, СуммаНачисления, СуммаФОТ, СуммаНалогов)
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


# Знайти документ за грудень 2025
q = erp.NewObject("Запрос")
q.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка "
    "ИЗ Документ.А_ОтражениеЗПпоКазне КАК Д "
    "ГДЕ Д.Номер = &Н И Д.Дата МЕЖДУ &Д1 И &Д2"
)
q.SetParameter("Н", "000000001")
q.SetParameter("Д1", com_date(2025, 12, 31, 0, 0, 0))
q.SetParameter("Д2", com_date(2025, 12, 31, 23, 59, 59))
sel = q.Execute().Выбрать()
if not sel.Следующий():
    print("ПОМИЛКА: документ не знайдено")
    sys.exit(1)

obj = sel.Ссылка.ПолучитьОбъект()
print(f"Документ: {S(obj.Ссылка)}")
print(f"Дата: {S(obj.Дата)}, Організация: {S(obj.Организация)}")

# Виклик завантаження (спровокує нову агрегацію)
print("\n>>> Виклик ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП() ...")
try:
    obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
    print(">>> OK")
except Exception as e:
    print(f">>> ПОМИЛКА: {e}")
    sys.exit(1)

# Перевірки
kz = obj.РаспределениеКазна
n = kz.Количество()
print(f"\nРаспределениеКазна: {n} рядків")

total_sum = 0.0
total_nach = 0.0
total_fot = 0.0
total_nal = 0.0
rows_with_nach = 0
rows_with_fot = 0
rows_with_nal = 0
rows_sum_mismatch = 0  # Сумма != СуммаНач+ФОТ+Нал (з толерантністю 0.01)

for i in range(n):
    r = kz.Получить(i)
    s = parse_num(r.Сумма)
    sn = parse_num(r.СуммаНачисления)
    sf = parse_num(r.СуммаФОТ)
    sna = parse_num(r.СуммаНалогов)
    total_sum += s
    total_nach += sn
    total_fot += sf
    total_nal += sna
    if sn != 0:
        rows_with_nach += 1
    if sf != 0:
        rows_with_fot += 1
    if sna != 0:
        rows_with_nal += 1
    expected = sn + sf + sna
    if abs(expected - s) > 0.01:
        rows_sum_mismatch += 1

print(f"\n=== Тотали ===")
print(f"  Сумма:           {total_sum:>15,.2f}")
print(f"  СуммаНачисления: {total_nach:>15,.2f}")
print(f"  СуммаФОТ:        {total_fot:>15,.2f}")
print(f"  СуммаНалогов:    {total_nal:>15,.2f}")
print(f"  Σ(Нач+ФОТ+Нал):  {total_nach + total_fot + total_nal:>15,.2f}")
print(f"  Різниця Сумма - Σ: {(total_sum - (total_nach + total_fot + total_nal)):,.2f}")

print(f"\n=== Розподіл по рядках ===")
print(f"  Рядків з СуммаНачисления <> 0: {rows_with_nach}")
print(f"  Рядків з СуммаФОТ <> 0:        {rows_with_fot}")
print(f"  Рядків з СуммаНалогов <> 0:    {rows_with_nal}")
print(f"  Рядків де Сумма ≠ Σ(Нач+ФОТ+Нал): {rows_sum_mismatch}")

# Зразки
print(f"\n=== Зразки перших 5 рядків ===")
for i in range(min(5, n)):
    r = kz.Получить(i)
    print(f"  [{i+1}] {S(r.ФИО)[:35]:35s}  Сум={parse_num(r.Сумма):>10.2f}  Нач={parse_num(r.СуммаНачисления):>10.2f}  ФОТ={parse_num(r.СуммаФОТ):>10.2f}  Нал={parse_num(r.СуммаНалогов):>10.2f}")

# Перевірки
print(f"\n=== ПЕРЕВІРКИ ===")
checks = [
    (n > 0, "РаспределениеКазна не пуста"),
    (total_sum > 0, "Сумма тотал > 0"),
    (rows_with_nach > 0, "Є рядки з СуммаНачисления"),
    (rows_sum_mismatch < n // 10, f"Більшість рядків мають Сумма ≈ Σ (mismatch={rows_sum_mismatch}/{n})"),
]
all_ok = True
for ok, desc in checks:
    mark = "OK" if ok else "FAIL"
    print(f"  [{mark}] {desc}")
    if not ok:
        all_ok = False

print()
if all_ok:
    print("==> ЧАСТИНА A (PIVOT) ПРОЙДЕНА")
    sys.exit(0)
else:
    print("==> ПОМИЛКИ")
    sys.exit(1)
