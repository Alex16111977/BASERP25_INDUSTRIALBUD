# -*- coding: utf-8 -*-
"""
Полная регрессия РаспределениеФ2 контрольных сумм/флагов по ВСЕМ документам (одна COM-сессия).
Для каждого документа: читает источники из объекта, считает ожидаемое (модель v4),
записывает документ (триггер ПередЗаписью), читает сохранённое ИЗ ТОГО ЖЕ объекта, сверяет.

Модель v4 (зеркало BSL _ЗаполнитьКонтрольныеСуммы):
  - карты Σ по сотруднику, пустой Сотрудник исключён;
  - колонки: итог по сотруднику в ПЕРВОЙ строке (ТЧ-порядок), дубли/пустые -> 0;
  - построчные флаги Сотрудники.А_РасхождениеРаспределение / А_РасхождениеДДС — по АГРЕГАТУ сотрудника;
  - общие флаги шапки: А_Расхождение = распр-расхождение по объединению ключей,
    А_РасхождениеОбщееДДС = ддс-расхождение (при Количество(ДДС)>0).
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

q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ Ссылка КАК Ссылка, Номер КАК Номер ИЗ Документ.РаспределениеФ2 КАК Д УПОРЯДОЧИТЬ ПО Д.Номер"
tz = q.Execute().Выгрузить()
refs = [(tz.Получить(i).Номер, tz.Получить(i).Ссылка) for i in range(tz.Количество())]
print(f"Документов РаспределениеФ2: {len(refs)}")

всего = 0
passed = 0
fails = []
hdr_rasp_T = 0
hdr_dds_T = 0

for номер, ref in refs:
    всего += 1
    obj = ref.ПолучитьОбъект()

    sotr = [(uid(obj.Сотрудники[i].Сотрудник), float(obj.Сотрудники[i].СуммаНачисления))
            for i in range(obj.Сотрудники.Количество())]
    rasp = [(uid(obj.Распределение[i].Сотрудник), float(obj.Распределение[i].СуммаНачисления))
            for i in range(obj.Распределение.Количество())]
    dds = [(uid(obj.А_ДвижениеДДС[i].Сотрудник), float(obj.А_ДвижениеДДС[i].Сумма))
           for i in range(obj.А_ДвижениеДДС.Количество())]

    map_rasp, map_dds, map_nach = {}, {}, {}
    for u, s in rasp:
        if u != ZERO:
            map_rasp[u] = map_rasp.get(u, 0.0) + s
    for u, s in dds:
        if u != ZERO:
            map_dds[u] = map_dds.get(u, 0.0) + s
    for u, s in sotr:
        if u != ZERO:
            map_nach[u] = map_nach.get(u, 0.0) + s

    есть_ддс = len(dds) > 0

    # ожидаемые колонки + построчные флаги (по агрегату сотрудника)
    exp_cols, exp_row_rasp, exp_row_dds = [], [], []
    seen = set()
    for u, n in sotr:
        if u != ZERO:
            inn = map_nach.get(u, 0.0)
            ir = map_rasp.get(u, 0.0)
            idd = map_dds.get(u, 0.0)
            exp_row_rasp.append(r2(inn - ir) != 0)
            exp_row_dds.append(есть_ддс and (r2(inn - idd) != 0))
            if u not in seen:
                seen.add(u)
                exp_cols.append((ir, idd))
            else:
                exp_cols.append((0.0, 0.0))
        else:
            exp_row_rasp.append(False)
            exp_row_dds.append(False)
            exp_cols.append((0.0, 0.0))

    # ожидаемые общие флаги шапки (объединение ключей)
    all_keys = set(map_nach) | set(map_rasp) | set(map_dds)
    exp_hdr_rasp = False
    exp_hdr_dds = False
    for u in all_keys:
        inn = map_nach.get(u, 0.0)
        ir = map_rasp.get(u, 0.0)
        idd = map_dds.get(u, 0.0)
        if r2(inn - ir) != 0:
            exp_hdr_rasp = True
        if есть_ддс and r2(inn - idd) != 0:
            exp_hdr_dds = True

    try:
        obj.Записать()
    except Exception as e:
        fails.append((номер, f"WRITE EXCEPTION: {e}"))
        continue

    # сохранённое из того же объекта
    bad = 0
    for i in range(obj.Сотрудники.Количество()):
        row = obj.Сотрудники[i]
        er, ed = exp_cols[i]
        if (r2(float(row.СуммаПоРаспределение) - er) != 0
                or r2(float(row.СуммаПоДДС) - ed) != 0
                or bool(row.А_РасхождениеРаспределение) != exp_row_rasp[i]
                or bool(row.А_РасхождениеДДС) != exp_row_dds[i]):
            bad += 1
    hdr_rasp = bool(obj.А_Расхождение)
    hdr_dds = bool(obj.А_РасхождениеОбщееДДС)
    if hdr_rasp:
        hdr_rasp_T += 1
    if hdr_dds:
        hdr_dds_T += 1

    if bad == 0 and hdr_rasp == exp_hdr_rasp and hdr_dds == exp_hdr_dds:
        passed += 1
    else:
        fails.append((номер, f"строк-расх={bad} | А_Расхождение={hdr_rasp}(ож {exp_hdr_rasp}) "
                             f"А_РасхождениеОбщееДДС={hdr_dds}(ож {exp_hdr_dds})"))

print(f"\nИтого: {всего}  PASS={passed}  FAIL={len(fails)}")
print(f"Шапка: А_Расхождение(распр)=True у {hdr_rasp_T}  А_РасхождениеОбщееДДС=True у {hdr_dds_T}")
if fails:
    print("\n--- FAILS ---")
    for номер, msg in fails:
        print(f"  {номер}: {msg}")
    sys.exit(1)
else:
    print(f"\n=== ALL PASS ({passed}/{всего}) ===")
    sys.exit(0)
