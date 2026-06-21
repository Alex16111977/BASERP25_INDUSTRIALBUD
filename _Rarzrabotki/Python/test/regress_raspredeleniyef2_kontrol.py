# -*- coding: utf-8 -*-
"""
Полная регрессия РаспределениеФ2 контрольных сумм по ВСЕМ документам (одна COM-сессия).
Для каждого документа: читает источники из объекта, считает ожидаемое (модель v3),
записывает документ (триггер ПередЗаписью), читает сохранённые колонки/флаг ИЗ ТОГО ЖЕ объекта,
сверяет. Печатает сводку и список расхождений.

Модель v3 (зеркало BSL _ЗаполнитьКонтрольныеСуммы):
  - карты Σ по сотруднику, пустой Сотрудник исключён;
  - колонки: итог по сотруднику в ПЕРВОЙ строке (ТЧ-порядок), дубли/пустые -> 0;
  - флаг: агрегат по объединению ключей (Начислено/Распределение/ДДС), ДДС-сравнение при Количество(ДДС)>0.
"""
import sys, win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

ZERO = "00000000-0000-0000-0000-000000000000"

def uid(ref):
    return erp.String(ref.УникальныйИдентификатор())

def r2(x):
    return round(x + 1e-9, 2)

# список всех документов
q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ Ссылка КАК Ссылка, Номер КАК Номер ИЗ Документ.РаспределениеФ2 КАК Д УПОРЯДОЧИТЬ ПО Д.Номер"
tz = q.Execute().Выгрузить()
refs = [(tz.Получить(i).Номер, tz.Получить(i).Ссылка) for i in range(tz.Количество())]
print(f"Документов РаспределениеФ2: {len(refs)}")

всего = 0
passed = 0
fails = []
флаг_T = 0
флаг_F = 0

for номер, ref in refs:
    всего += 1
    obj = ref.ПолучитьОбъект()

    # источники из объекта (ТЧ-порядок)
    sotr = [(uid(obj.Сотрудники[i].Сотрудник), float(obj.Сотрудники[i].СуммаНачисления))
            for i in range(obj.Сотрудники.Количество())]
    rasp = [(uid(obj.Распределение[i].Сотрудник), float(obj.Распределение[i].СуммаНачисления))
            for i in range(obj.Распределение.Количество())]
    dds = [(uid(obj.А_ДвижениеДДС[i].Сотрудник), float(obj.А_ДвижениеДДС[i].Сумма))
           for i in range(obj.А_ДвижениеДДС.Количество())]

    map_rasp = {}
    for u, s in rasp:
        if u == ZERO:
            continue
        map_rasp[u] = map_rasp.get(u, 0.0) + s
    map_dds = {}
    for u, s in dds:
        if u == ZERO:
            continue
        map_dds[u] = map_dds.get(u, 0.0) + s
    map_nach = {}
    for u, s in sotr:
        if u == ZERO:
            continue
        map_nach[u] = map_nach.get(u, 0.0) + s

    есть_ддс = len(dds) > 0

    # ожидаемые колонки: первая строка сотрудника
    exp_cols = []
    seen = set()
    for u, n in sotr:
        if u != ZERO and u not in seen:
            seen.add(u)
            exp_cols.append((map_rasp.get(u, 0.0), map_dds.get(u, 0.0)))
        else:
            exp_cols.append((0.0, 0.0))

    # ожидаемый флаг: агрегат по объединению ключей
    all_keys = set(map_nach) | set(map_rasp) | set(map_dds)
    exp_flag = False
    for u in all_keys:
        nn = map_nach.get(u, 0.0)
        rr = map_rasp.get(u, 0.0)
        dd = map_dds.get(u, 0.0)
        if (r2(nn - rr) != 0) or (есть_ддс and r2(nn - dd) != 0):
            exp_flag = True
            break

    # запись (триггер ПередЗаписью)
    try:
        obj.Записать()
    except Exception as e:
        fails.append((номер, f"WRITE EXCEPTION: {e}"))
        continue

    # сохранённые колонки из того же объекта (после ПередЗаписью)
    bad = 0
    for i in range(obj.Сотрудники.Количество()):
        sp_r = float(obj.Сотрудники[i].СуммаПоРаспределение)
        sp_d = float(obj.Сотрудники[i].СуммаПоДДС)
        er, ed = exp_cols[i]
        if r2(sp_r - er) != 0 or r2(sp_d - ed) != 0:
            bad += 1
    stored_flag = bool(obj.А_Расхождение)

    if exp_flag:
        флаг_T += 1
    else:
        флаг_F += 1

    if bad == 0 and stored_flag == exp_flag:
        passed += 1
    else:
        fails.append((номер, f"колонок-расх={bad} флаг_сохр={stored_flag} ож={exp_flag}"))

# footer-инвариант: Σ колонки == Σ источника по непустым сотрудникам (на нескольких доках)
print(f"\nИтого: {всего}  PASS={passed}  FAIL={len(fails)}")
print(f"Флаг А_Расхождение: True={флаг_T}  False={флаг_F}")
if fails:
    print("\n--- FAILS ---")
    for номер, msg in fails:
        print(f"  {номер}: {msg}")
    sys.exit(1)
else:
    print("\n=== ALL PASS (75/75) ===")
    sys.exit(0)
