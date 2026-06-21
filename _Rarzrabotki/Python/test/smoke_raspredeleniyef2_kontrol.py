# -*- coding: utf-8 -*-
"""
Smoke РаспределениеФ2: контрольные суммы Сотрудники.СуммаПоРаспределение / СуммаПоДДС + А_Расхождение.

Режимы:
  python smoke_raspredeleniyef2_kontrol.py [НОМЕР]          -> только ОЖИДАЕМЫЕ суммы (без записи)
  python smoke_raspredeleniyef2_kontrol.py [НОМЕР] write    -> записать документ и сверить СОХРАНЁННЫЕ колонки с ожидаемыми

Ожидаемое (эталон 000000082): 44 строки Начисл==Распр==ДДС; 3 битых строки Распр==Начисл, ДДС==0; А_Расхождение=Истина.
"""
import sys, win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

NOMER = "000000082"
DO_WRITE = False
for a in sys.argv[1:]:
    if a.lower() == "write":
        DO_WRITE = True
    else:
        NOMER = a

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

ZERO = "00000000-0000-0000-0000-000000000000"

def uid(ref):
    u = ref.УникальныйИдентификатор()
    try:
        return erp.String(u)
    except Exception:
        return erp.string(u)

def r2(x):
    return round(x + 1e-9, 2)

def read_rows(tabular, sumfield):
    q = erp.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ Сотрудник КАК Сотрудник, " + sumfield + " КАК Сум "
              "ИЗ Документ.РаспределениеФ2." + tabular + " КАК Т ГДЕ Т.Ссылка.Номер = &Н "
              "УПОРЯДОЧИТЬ ПО Т.НомерСтроки")
    q.SetParameter("Н", NOMER)
    tz = q.Execute().Выгрузить()
    rows = []
    for i in range(tz.Количество()):
        r = tz.Получить(i)
        rows.append((uid(r.Сотрудник), float(r.Сум)))
    return rows

rasp_rows = read_rows("Распределение", "СуммаНачисления")
dds_rows  = read_rows("А_ДвижениеДДС", "Сумма")
sotr_rows = read_rows("Сотрудники", "СуммаНачисления")

map_rasp = {}
for u, s in rasp_rows:
    if u == ZERO:
        continue
    map_rasp[u] = map_rasp.get(u, 0.0) + s
map_dds = {}
for u, s in dds_rows:
    if u == ZERO:
        continue
    map_dds[u] = map_dds.get(u, 0.0) + s

есть_ддс = len(dds_rows) > 0

print(f"=== Документ №{NOMER} ===")
print(f"строк Сотрудники={len(sotr_rows)}  Распределение={len(rasp_rows)}  А_ДвижениеДДС={len(dds_rows)}")

# карта «начислено по сотруднику» (для агрегатного флага); пустой Сотрудник исключаем
map_nach = {}
for u, s in sotr_rows:
    if u == ZERO:
        continue
    map_nach[u] = map_nach.get(u, 0.0) + s

# --- ОЖИДАЕМЫЕ колонки: «итог по сотруднику» в ПЕРВОЙ строке сотрудника; дубли/пустые -> 0 ---
expected = {}
seen = set()
for idx, (u, начисл) in enumerate(sotr_rows):
    if u != ZERO and u not in seen:
        seen.add(u)
        er = map_rasp.get(u, 0.0)
        ed = map_dds.get(u, 0.0)
    else:
        er = 0.0
        ed = 0.0
    expected[(idx, u)] = (начисл, er, ed)

# --- ОЖИДАЕМЫЙ флаг А_Расхождение: агрегат по сотруднику (объединение ключей; пустой исключён) ---
all_keys = set(map_nach) | set(map_rasp) | set(map_dds)
есть_расхождение = False
расх_сотр = 0
for u in all_keys:
    н = map_nach.get(u, 0.0)
    р = map_rasp.get(u, 0.0)
    д = map_dds.get(u, 0.0)
    if (r2(н - р) != 0) or (есть_ддс and r2(н - д) != 0):
        есть_расхождение = True
        расх_сотр += 1

green_rows = sum(1 for u, n in sotr_rows
                 if u != ZERO and r2(n - map_rasp.get(u, 0.0)) == 0
                 and (not есть_ддс or r2(n - map_dds.get(u, 0.0)) == 0))
print(f"ОЖИДАЕМО: строк-green={green_rows}/{len(sotr_rows)}  сотрудников-с-расхождением={расх_сотр}  А_Расхождение={есть_расхождение}")

if not DO_WRITE:
    print("\n[режим: только ожидаемые суммы; запись не делалась]")
    sys.exit(0)

# --- ЗАПИСЬ документа (триггер ПередЗаписью) ---
q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК Ссылка ИЗ Документ.РаспределениеФ2 КАК Д ГДЕ Д.Номер = &Н"
q.SetParameter("Н", NOMER)
tz = q.Execute().Выгрузить()
ref = tz.Получить(0).Ссылка
obj = ref.ПолучитьОбъект()
было_проведен = bool(obj.Проведен)
try:
    obj.Записать()
except Exception:
    obj.Write()
print(f"\n[записан; был проведён={было_проведен}, стал проведён={bool(ref.ПолучитьОбъект().Проведен)}]")

# --- читаем СОХРАНЁННЫЕ колонки ---
q = erp.NewObject("Запрос")
q.Text = ("ВЫБРАТЬ Сотрудник КАК Сотрудник, СуммаНачисления КАК Нач, "
          "СуммаПоРаспределение КАК Распр, СуммаПоДДС КАК ДДС "
          "ИЗ Документ.РаспределениеФ2.Сотрудники КАК Т ГДЕ Т.Ссылка.Номер = &Н "
          "УПОРЯДОЧИТЬ ПО Т.НомерСтроки")
q.SetParameter("Н", NOMER)
tz = q.Execute().Выгрузить()
stored = []
for i in range(tz.Количество()):
    r = tz.Получить(i)
    stored.append((uid(r.Сотрудник), float(r.Нач), float(r.Распр), float(r.ДДС)))

# А_Расхождение из шапки
q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ А_Расхождение КАК Расх ИЗ Документ.РаспределениеФ2 КАК Д ГДЕ Д.Номер = &Н"
q.SetParameter("Н", NOMER)
tz = q.Execute().Выгрузить()
храним_расх = bool(tz.Получить(0).Расх)

# --- сверка построчно (по индексу строки) ---
fail = 0
for idx, ((u_s, нач_s, распр_s, ддс_s)) in enumerate(stored):
    нач_e, распр_e, ддс_e = expected[(idx, u_s)]
    if r2(распр_s - распр_e) != 0 or r2(ддс_s - ддс_e) != 0:
        fail += 1
        if fail <= 10:
            print(f"  FAIL стр{idx}: Распр сохр={распр_s} ож={распр_e} | ДДС сохр={ддс_s} ож={ддс_e}")

print(f"\nСверка СОХРАНЁННЫХ колонок: строк={len(stored)} расхождений с ожидаемым={fail}")
print(f"А_Расхождение сохранён={храним_расх} (ожидаемо={есть_расхождение})")

ok = (fail == 0) and (храним_расх == есть_расхождение)
print("\n=== " + ("PASS" if ok else "FAIL") + " ===")
sys.exit(0 if ok else 1)
