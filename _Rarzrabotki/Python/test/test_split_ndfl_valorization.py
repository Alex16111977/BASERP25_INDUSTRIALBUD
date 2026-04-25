# -*- coding: utf-8 -*-
# Тест: НДФЛ/ВС идёт отдельной строкой в НачисленнаяЗарплатаИВзносыПоФизлицам
# (а не склеивается с NET в валоризации gross)
#
# Проверяем на документе А_ОтражениеЗПпоКазне 000000001 від 31.12.2025:
# - Постернак: NET = 100 000 на "Зарплата управленческого персонала"
#              НДФЛ = 8 337,81 отдельно на "Удержание с админ зарплату"
# - Инвариант: SUM(Сумма) по ФЛ = NET + НДФЛ (общая сумма не изменилась)
# - НачисленныйНДФЛ — без изменений (Op 3 не менялся)
import win32com.client, sys, pywintypes, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String


def parse_num(val):
    s = S(val).replace('\xa0', '').replace(' ', '').replace(',', '.')
    return float(s or "0")


# === Знайти батьківський документ ===
q = erp.NewObject("Запрос")
q.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка "
    "ИЗ Документ.А_ОтражениеЗПпоКазне КАК Д "
    "ГДЕ Д.Номер = &Н"
)
q.SetParameter("Н", "000000001")
sel = q.Execute().Выбрать()
if not sel.Следующий():
    print("ПОМИЛКА: документ не знайдено")
    sys.exit(1)

doc_ref = sel.Ссылка
obj = doc_ref.ПолучитьОбъект()
print(f"Документ: {S(doc_ref)}")

# === Виклик заповнення ===
print("\n>>> ЗаполнитьДляОтражения_ДокументОтражениеЗарплатыВФинансовомУчете()...")
try:
    obj.ЗаполнитьДляОтражения_ДокументОтражениеЗарплатыВФинансовомУчете()
    print(">>> OK")
except Exception as e:
    print(f">>> ПОМИЛКА: {e}")
    sys.exit(1)

# === Збираємо дані для Постернака ===
print("\n=== Постернак Андрій — НачисленнаяЗарплатаИВзносыПоФизлицам ===")
posternak_rows = []
posternak_fl = None
for i in range(obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Количество()):
    r = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Получить(i)
    name = S(r.ФизическоеЛицо.Наименование) if erp.ЗначениеЗаполнено(r.ФизическоеЛицо) else ""
    if "Постернак" in name:
        sposob = S(r.СпособОтраженияЗарплатыВБухучете) if erp.ЗначениеЗаполнено(r.СпособОтраженияЗарплатыВБухучете) else "<пусто>"
        summa = parse_num(r.Сумма)
        vznosy = parse_num(r.ВзносыВсего)
        posternak_rows.append({"sposob": sposob, "summa": summa, "vznosy": vznosy})
        posternak_fl = r.ФизическоеЛицо
        print(f"  [{i+1}] Спосіб={sposob[:45]:45s} Сума={summa:>12.2f}  Внески={vznosy:>10.2f}")

# Сумарно по Постернаку
total_summa = sum(r["summa"] for r in posternak_rows)
total_vznosy = sum(r["vznosy"] for r in posternak_rows)
print(f"  ВСЬОГО: Сума={total_summa:.2f}, Внески={total_vznosy:.2f}")

# Перевірки для Постернака
zp_row = [r for r in posternak_rows if "управленческ" in r["sposob"].lower()]
ndfl_row = [r for r in posternak_rows if "удержан" in r["sposob"].lower()]
fot_row = [r for r in posternak_rows if "начислен" in r["sposob"].lower() and r["vznosy"] > 0]

zp_summa = zp_row[0]["summa"] if zp_row else 0
ndfl_summa = ndfl_row[0]["summa"] if ndfl_row else 0
fot_vznosy = fot_row[0]["vznosy"] if fot_row else 0

# === НачисленныйНДФЛ для Постернака ===
print(f"\n=== Постернак — НачисленныйНДФЛ ===")
ndfl_total = 0.0
ndfl_cnt = 0
for i in range(obj.НачисленныйНДФЛ.Количество()):
    r = obj.НачисленныйНДФЛ.Получить(i)
    name = S(r.ФизическоеЛицо.Наименование) if erp.ЗначениеЗаполнено(r.ФизическоеЛицо) else ""
    if "Постернак" in name:
        s = parse_num(r.Сумма)
        ndfl_total += s
        ndfl_cnt += 1
        print(f"  [{ndfl_cnt}] Тип={S(r.ТипНалога):20s} Сума={s:>10.2f}")
print(f"  ВСЬОГО НДФЛ+ВС: {ndfl_total:.2f}")

# === Глобальний інваріант: SUM всіх ФЛ ===
print(f"\n=== Глобальні суми ===")
total_all_summa = 0.0
total_all_vznosy = 0.0
for i in range(obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Количество()):
    r = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Получить(i)
    total_all_summa += parse_num(r.Сумма)
    total_all_vznosy += parse_num(r.ВзносыВсего)
print(f"  НачисленнаяЗарплатаИВзносыПоФизлицам: Сума={total_all_summa:.2f}, Внески={total_all_vznosy:.2f}")
print(f"  Рядків: {obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Количество()}")

# === ПЕРЕВІРКИ ===
print("\n=== ПЕРЕВІРКИ ===")
checks = [
    (len(posternak_rows) >= 3, f"Постернак: {len(posternak_rows)} рядків (≥3: NET + НДФЛ + ФОТ)"),
    (len(zp_row) > 0, "Є рядок 'Зарплата управленческого персонала'"),
    (abs(zp_summa - 100000.0) < 0.01,
     f"Зарплата управленческого = {zp_summa:.2f} (очікується 100 000,00 — чистий NET)"),
    (len(ndfl_row) > 0, "Є рядок 'Удержание с админ зарплату'"),
    (abs(ndfl_summa - 8337.81) < 0.01,
     f"Удержание с админ = {ndfl_summa:.2f} (очікується 8 337,81 — НДФЛ окремо)"),
    (len(fot_row) > 0, "Є рядок ФОТ 'Начисления на админ зарплату'"),
    (abs(fot_vznosy - 7975.29) < 0.01,
     f"ФОТ Внески = {fot_vznosy:.2f} (очікується 7 975,29 — без змін)"),
    (abs(total_summa - (100000.0 + 8337.81)) < 0.01,
     f"Інваріант Постернак: SUM(Сума)={total_summa:.2f} == NET+НДФЛ={100000+8337.81:.2f}"),
    (abs(ndfl_total - 8337.81) < 0.01,
     f"НачисленныйНДФЛ Постернак: {ndfl_total:.2f} == 8337.81 (Op 3 без змін)"),
    (ndfl_cnt == 2, f"НачисленныйНДФЛ: {ndfl_cnt} рядків (НДФЛ + ВС)"),
]

all_ok = True
for ok, desc in checks:
    mark = "OK" if ok else "FAIL"
    print(f"  [{mark}] {desc}")
    if not ok:
        all_ok = False

print()
if all_ok:
    print("==> ТЕСТ РОЗДІЛЕННЯ ВАЛОРИЗАЦІЇ НДФЛ ПРОЙДЕНО")
    sys.exit(0)
else:
    print("==> ПОМИЛКИ")
    sys.exit(1)
