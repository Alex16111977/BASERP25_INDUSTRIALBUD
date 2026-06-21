# -*- coding: utf-8 -*-
"""
Smoke РаспределениеФ2: контрольные суммы/флаги (v4).
  Сотрудники.СуммаПоРаспределение / СуммаПоДДС — итог по сотруднику в первой строке;
  Сотрудники.А_РасхождениеРаспределение / А_РасхождениеДДС — построчно (по агрегату сотрудника);
  шапка А_Расхождение — общий по распределению; А_РасхождениеОбщееДДС — общий по ДДС.

Режимы:
  python smoke_raspredeleniyef2_kontrol.py [НОМЕР]         -> ожидаемое (без записи)
  python smoke_raspredeleniyef2_kontrol.py [НОМЕР] write   -> записать и сверить сохранённое с ожидаемым
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
    return erp.String(ref.УникальныйИдентификатор())

def r2(x):
    return round(x + 1e-9, 2)

def read_rows(tabular, sumfield):
    q = erp.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ Сотрудник КАК Сотрудник, " + sumfield + " КАК Сум "
              "ИЗ Документ.РаспределениеФ2." + tabular + " КАК Т ГДЕ Т.Ссылка.Номер = &Н "
              "УПОРЯДОЧИТЬ ПО Т.НомерСтроки")
    q.SetParameter("Н", NOMER)
    tz = q.Execute().Выгрузить()
    return [(uid(tz.Получить(i).Сотрудник), float(tz.Получить(i).Сум)) for i in range(tz.Количество())]

rasp_rows = read_rows("Распределение", "СуммаНачисления")
dds_rows = read_rows("А_ДвижениеДДС", "Сумма")
sotr_rows = read_rows("Сотрудники", "СуммаНачисления")

map_rasp, map_dds, map_nach = {}, {}, {}
for u, s in rasp_rows:
    if u != ZERO:
        map_rasp[u] = map_rasp.get(u, 0.0) + s
for u, s in dds_rows:
    if u != ZERO:
        map_dds[u] = map_dds.get(u, 0.0) + s
for u, s in sotr_rows:
    if u != ZERO:
        map_nach[u] = map_nach.get(u, 0.0) + s

есть_ддс = len(dds_rows) > 0

print(f"=== Документ №{NOMER} ===")
print(f"строк Сотрудники={len(sotr_rows)}  Распределение={len(rasp_rows)}  А_ДвижениеДДС={len(dds_rows)}")

# ожидаемые колонки + построчные флаги (по агрегату сотрудника)
exp = {}  # idx -> (распр, ддс, фл_распр, фл_ддс)
seen = set()
for idx, (u, n) in enumerate(sotr_rows):
    if u != ZERO:
        inn = map_nach.get(u, 0.0)
        ir = map_rasp.get(u, 0.0)
        idd = map_dds.get(u, 0.0)
        fr = r2(inn - ir) != 0
        fd = есть_ддс and (r2(inn - idd) != 0)
        if u not in seen:
            seen.add(u)
            exp[idx] = (ir, idd, fr, fd)
        else:
            exp[idx] = (0.0, 0.0, fr, fd)
    else:
        exp[idx] = (0.0, 0.0, False, False)

# ожидаемые общие флаги шапки
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

стр_распр = sum(1 for v in exp.values() if v[2])
стр_ддс = sum(1 for v in exp.values() if v[3])
print(f"ОЖИДАЕМО: строк с А_РасхождениеРаспределение={стр_распр}  А_РасхождениеДДС={стр_ддс}")
print(f"          шапка А_Расхождение(распр)={exp_hdr_rasp}  А_РасхождениеОбщееДДС={exp_hdr_dds}")

if not DO_WRITE:
    print("\n[режим: только ожидаемое; запись не делалась]")
    sys.exit(0)

# --- запись (триггер ПередЗаписью) ---
q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК Ссылка ИЗ Документ.РаспределениеФ2 КАК Д ГДЕ Д.Номер = &Н"
q.SetParameter("Н", NOMER)
ref = q.Execute().Выгрузить().Получить(0).Ссылка
obj = ref.ПолучитьОбъект()
было = bool(obj.Проведен)
obj.Записать()
print(f"\n[записан; был проведён={было}, стал={bool(ref.ПолучитьОбъект().Проведен)}]")

# --- сохранённые строки ---
q = erp.NewObject("Запрос")
q.Text = ("ВЫБРАТЬ Сотрудник КАК Сотрудник, СуммаПоРаспределение КАК Распр, СуммаПоДДС КАК ДДС, "
          "А_РасхождениеРаспределение КАК ФлР, А_РасхождениеДДС КАК ФлД "
          "ИЗ Документ.РаспределениеФ2.Сотрудники КАК Т ГДЕ Т.Ссылка.Номер = &Н "
          "УПОРЯДОЧИТЬ ПО Т.НомерСтроки")
q.SetParameter("Н", NOMER)
tz = q.Execute().Выгрузить()
fail = 0
for i in range(tz.Количество()):
    r = tz.Получить(i)
    er, ed, fr, fd = exp[i]
    if (r2(float(r.Распр) - er) != 0 or r2(float(r.ДДС) - ed) != 0
            or bool(r.ФлР) != fr or bool(r.ФлД) != fd):
        fail += 1
        if fail <= 10:
            print(f"  FAIL стр{i}: Распр={float(r.Распр)}/{er} ДДС={float(r.ДДС)}/{ed} "
                  f"ФлР={bool(r.ФлР)}/{fr} ФлД={bool(r.ФлД)}/{fd}")

q = erp.NewObject("Запрос")
q.Text = ("ВЫБРАТЬ А_Расхождение КАК Р, А_РасхождениеОбщееДДС КАК Д "
          "ИЗ Документ.РаспределениеФ2 КАК Д ГДЕ Д.Номер = &Н")
q.SetParameter("Н", NOMER)
h = q.Execute().Выгрузить().Получить(0)
hdr_rasp, hdr_dds = bool(h.Р), bool(h.Д)

print(f"\nСтрок={tz.Количество()} расхождений с ожидаемым={fail}")
print(f"Шапка: А_Расхождение={hdr_rasp}(ож {exp_hdr_rasp})  А_РасхождениеОбщееДДС={hdr_dds}(ож {exp_hdr_dds})")
ok = (fail == 0) and (hdr_rasp == exp_hdr_rasp) and (hdr_dds == exp_hdr_dds)
print("\n=== " + ("PASS" if ok else "FAIL") + " ===")
sys.exit(0 if ok else 1)
