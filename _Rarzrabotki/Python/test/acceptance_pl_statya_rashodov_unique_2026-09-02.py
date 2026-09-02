# -*- coding: utf-8 -*-
"""Приёмка правила 5 справочника А_Статьи_PL: уникальность СтатьяРасходов между PL-статьями.

Все кейсы выполняются в транзакции с ОтменитьТранзакцию — база после прогона без следов.
Число проверок — константа 6 (ни одна проверка не стоит под if).
Кейсы:
  1. Базовая запись PL_X без изменений — без исключения.
  2. В PL_X добавлена ДДС_свободная (её А_СтатьяРасходов = СР_занятая с оборотами, уже в PL_Y) — исключение.
  3. Текст исключения кейса 2 содержит заголовок правила 5, СР_занятая, PL_Y и «Решение».
  4. В PL_X добавлена ДДС_занятая (стоит в PL_Z, есть обороты) — исключение правила 4 (ДДС), не правила 5.
  5. Shared без оборотов: одна СтатьяРасходов без оборотов оказывается в PL_W и PL_X — запись проходит
     (стратегии подбора A/B/C, см. блок «подбор»).
  6. База без следов: число строк ТЧ до и после прогона совпадает.
"""
import sys
import datetime
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ПРОВЕРОК = 0
ПРОВАЛОВ = 0


def проверка(name, cond, extra=""):
    global ПРОВЕРОК, ПРОВАЛОВ
    ПРОВЕРОК += 1
    if cond:
        print(f"  [OK]   {name}")
    else:
        ПРОВАЛОВ += 1
        print(f"  [FAIL] {name} {extra}")


def query(text, params=None):
    q = erp.NewObject("Запрос")
    q.Text = text
    for k, v in (params or {}).items():
        q.SetParameter(k, v)
    return q.Execute().Выгрузить()


def uid(ref):
    return S(ref.УникальныйИдентификатор())


def err_text(e):
    if hasattr(e, 'excepinfo') and e.excepinfo:
        return str(e.excepinfo[2])
    return str(e)


def write_in_transaction(actions):
    """actions: список функций, каждая делает Записать(). Возвращает (ok, text).
    Всегда откатывает транзакцию."""
    erp.НачатьТранзакцию()
    try:
        for a in actions:
            a()
        return True, ""
    except Exception as e:
        return False, err_text(e)
    finally:
        erp.ОтменитьТранзакцию()


# ------------------------------------------------------------------ период
now = datetime.datetime.now()
НП = datetime.datetime(now.year - 1, now.month, 1, 3, 0, 0)
КП = query("ВЫБРАТЬ КОНЕЦПЕРИОДА(&Д, МЕСЯЦ) КАК Д",
           {"Д": datetime.datetime(now.year, now.month, 1, 3, 0, 0)}).Получить(0).Д

# ------------------------------------------------------------------ данные
# обороты по СтатьяРасходов (uuid -> сумма)
r = query("""
ВЫБРАТЬ Об.СтатьяРасходов КАК СР, СУММА(Об.СуммаРеглПриход) КАК Сумма
ИЗ РегистрНакопления.ПрочиеРасходы.Обороты(&НП, &КП, , ) КАК Об
СГРУППИРОВАТЬ ПО Об.СтатьяРасходов
ИМЕЮЩИЕ СУММА(Об.СуммаРеглПриход) > 0
""", {"НП": НП, "КП": КП})
sr_turn = {}
for i in range(r.Количество()):
    row = r.Получить(i)
    sr_turn[uid(row.СР)] = float(row.Сумма)

# обороты по ДДС (через СтатьяРасходов.А_СтатьяДвиженияДенежныхСредств — как в правиле 4)
r = query("""
ВЫБРАТЬ Об.СтатьяРасходов.А_СтатьяДвиженияДенежныхСредств КАК ДДС, СУММА(Об.СуммаРеглПриход) КАК Сумма
ИЗ РегистрНакопления.ПрочиеРасходы.Обороты(&НП, &КП, , ) КАК Об
СГРУППИРОВАТЬ ПО Об.СтатьяРасходов.А_СтатьяДвиженияДенежныхСредств
ИМЕЮЩИЕ СУММА(Об.СуммаРеглПриход) > 0
""", {"НП": НП, "КП": КП})
dds_turn = set()
for i in range(r.Количество()):
    row = r.Получить(i)
    if not row.ДДС.Пустая():
        dds_turn.add(uid(row.ДДС))

# строки ТЧ всех PL (без пометки удаления)
r = query("""
ВЫБРАТЬ ТЧ.Ссылка КАК PL, ТЧ.Ссылка.Код КАК Код, ТЧ.СтатьяДвиженияДенежныхСредств КАК ДДС, ТЧ.СтатьяРасходов КАК СР
ИЗ Справочник.А_Статьи_PL.Статьи КАК ТЧ
ГДЕ НЕ ТЧ.Ссылка.ПометкаУдаления
""")
ТЧ_строк_до = r.Количество()
pl_ref = {}        # uid PL -> ref
pl_name = {}       # uid PL -> "код имя"
pl_srs = {}        # uid PL -> set(uid СР)
sr_pls = {}        # uid СР -> set(uid PL)
sr_name = {}       # uid СР -> имя
dds_in_pl = {}     # uid ДДС -> uid PL
dds_ref = {}       # uid ДДС -> ref
dds_name = {}
dds_sr = {}        # uid ДДС (в ТЧ) -> uid СР
for i in range(r.Количество()):
    row = r.Получить(i)
    p = uid(row.PL)
    pl_ref[p] = row.PL
    pl_name[p] = f"{S(row.Код)} {S(row.PL)}"
    d = uid(row.ДДС)
    dds_in_pl[d] = p
    dds_ref[d] = row.ДДС
    dds_name[d] = S(row.ДДС)
    if not row.СР.Пустая():
        s = uid(row.СР)
        sr_name[s] = S(row.СР)
        pl_srs.setdefault(p, set()).add(s)
        sr_pls.setdefault(s, set()).add(p)
        dds_sr[d] = s

# ДДС каталога с заполненной А_СтатьяРасходов, не входящие ни в одну PL (свободные)
r = query("""
ВЫБРАТЬ ДДС.Ссылка КАК ДДС, ДДС.А_СтатьяРасходов КАК СР
ИЗ Справочник.СтатьиДвиженияДенежныхСредств КАК ДДС
ГДЕ НЕ ДДС.ПометкаУдаления И НЕ ДДС.ЭтоГруппа
    И ДДС.А_СтатьяРасходов <> ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиРасходов.ПустаяСсылка)
    И НЕ ДДС.Ссылка В (ВЫБРАТЬ Т.СтатьяДвиженияДенежныхСредств ИЗ Справочник.А_Статьи_PL.Статьи КАК Т
                       ГДЕ НЕ Т.Ссылка.ПометкаУдаления)
""")
sr_free_dds = {}   # uid СР -> [uid ДДС]
for i in range(r.Количество()):
    row = r.Получить(i)
    s = uid(row.СР)
    d = uid(row.ДДС)
    sr_name.setdefault(s, S(row.СР))
    dds_ref[d] = row.ДДС
    dds_name[d] = S(row.ДДС)
    sr_free_dds.setdefault(s, []).append(d)

# PL, которые правило 5 блокирует уже сейчас
blocked = set()
for s, pls in sr_pls.items():
    if len(pls) > 1 and sr_turn.get(s, 0) > 0:
        blocked |= pls

# ------------------------------------------------------------------ диагностика
print("=== ДИАГНОСТИКА: СтатьяРасходов в ≥2 PL (окно %s — %s) ===" % (НП.date(), КП))
for s, pls in sorted(sr_pls.items(), key=lambda kv: -sr_turn.get(kv[0], 0)):
    if len(pls) > 1:
        print(f"  {sr_name[s]} | PL={len(pls)} | обороты 12м={sr_turn.get(s, 0):,.2f} | "
              + "; ".join(sorted(pl_name[p] for p in pls)))
print(f"  PL-статей, которые сейчас не запишутся по правилу 5: {len(blocked)}")
for p in sorted(blocked, key=lambda x: pl_name[x]):
    print(f"    {pl_name[p]}")

# ------------------------------------------------------------------ подбор
СР_занятая = None
for s, pls in sr_pls.items():
    if sr_turn.get(s, 0) > 0 and sr_free_dds.get(s):
        СР_занятая = s
        break
PL_Y = next(iter(sr_pls[СР_занятая])) if СР_занятая else None
ДДС_свободная = sr_free_dds[СР_занятая][0] if СР_занятая else None

PL_X = None
for p in pl_ref:
    if p in blocked:
        continue
    if СР_занятая and СР_занятая in pl_srs.get(p, set()):
        continue
    PL_X = p
    break

ДДС_занятая = None
PL_Z = None
for d, p in dds_in_pl.items():
    if p != PL_X and d in dds_turn:
        ДДС_занятая, PL_Z = d, p
        break

# Кейс 5 — три стратегии подбора shared-статьи без оборотов:
#  A) СР0 уже стоит в PL_W и есть свободная ДДС с этой СР0 → в PL_X добавляем свободную ДДС
#  B) в PL_W есть ДДС без оборотов, чья СР0 без оборотов → в PL_X добавляем ту же ДДС (shared по обоим правилам)
#  C) СР0 без оборотов с ≥2 свободными ДДС → PL_W получает первую, PL_X — вторую
СР_без_оборотов = None
PL_W = None
кейс5_действия = []   # [(uid PL, uid ДДС)] в порядке записи
стратегия5 = ""
for s, pls in sr_pls.items():
    if sr_turn.get(s, 0) == 0 and sr_free_dds.get(s):
        w = next((p for p in pls if p != PL_X and p not in blocked), None)
        if w:
            СР_без_оборотов, PL_W, стратегия5 = s, w, "A"
            кейс5_действия = [(PL_X, sr_free_dds[s][0])]
            break
if not СР_без_оборотов:
    for d, p in dds_in_pl.items():
        s = dds_sr.get(d)
        if p != PL_X and p not in blocked and d not in dds_turn and s and sr_turn.get(s, 0) == 0:
            СР_без_оборотов, PL_W, стратегия5 = s, p, "B"
            кейс5_действия = [(PL_X, d)]
            break
if not СР_без_оборотов:
    for s, lst in sr_free_dds.items():
        if sr_turn.get(s, 0) == 0 and len(lst) >= 2:
            w = next((p for p in pl_ref if p != PL_X and p not in blocked), None)
            if w:
                СР_без_оборотов, PL_W, стратегия5 = s, w, "C"
                кейс5_действия = [(w, lst[0]), (PL_X, lst[1])]
                break

print("\n=== ПОДБОР ДАННЫХ ===")
print(f"  СР_занятая    : {sr_name.get(СР_занятая, '—')} (обороты {sr_turn.get(СР_занятая, 0):,.2f})")
print(f"  PL_Y          : {pl_name.get(PL_Y, '—')}")
print(f"  ДДС_свободная : {dds_name.get(ДДС_свободная, '—')}")
print(f"  PL_X          : {pl_name.get(PL_X, '—')}")
print(f"  ДДС_занятая   : {dds_name.get(ДДС_занятая, '—')} (в {pl_name.get(PL_Z, '—')})")
print(f"  СР_без_оборотов: {sr_name.get(СР_без_оборотов, '—')} (стратегия {стратегия5 or '—'}, "
      f"свободных ДДС {len(sr_free_dds.get(СР_без_оборотов, []))})")
print(f"  PL_W          : {pl_name.get(PL_W, '—')}")
for p, d in кейс5_действия:
    print(f"      кейс 5: в {pl_name.get(p, '—')} добавить ДДС «{dds_name.get(d, '—')}»")

# ------------------------------------------------------------------ кейсы
print("\n=== КЕЙСЫ ===")

# 1. базовая запись
def _c1():
    obj = pl_ref[PL_X].ПолучитьОбъект()
    obj.Записать()
ok1, t1 = write_in_transaction([_c1]) if PL_X else (False, "PL_X не подобран")
проверка("1. Базовая запись PL_X без изменений проходит", ok1, t1[:300])

# 2. правило 5
def _c2():
    obj = pl_ref[PL_X].ПолучитьОбъект()
    стр = obj.Статьи.Добавить()
    стр.СтатьяДвиженияДенежныхСредств = dds_ref[ДДС_свободная]
    obj.Записать()
ok2, t2 = write_in_transaction([_c2]) if (PL_X and ДДС_свободная) else (True, "данные не подобраны")
проверка("2. Добавление ДДС с занятой СтатьяРасходов (с оборотами) отказано", not ok2, t2[:300])
print("      текст отказа (как увидит финансист):")
for line in t2.strip().splitlines()[:8]:
    print("      | " + line.strip())

# 3. текст
маркеры = ["Статья расходов не может использоваться",
           sr_name.get(СР_занятая, "§НЕТ§"),
           S(pl_ref[PL_Y]) if PL_Y else "§НЕТ§",
           "Решение"]
проверка("3. Текст исключения: заголовок правила 5 + СР + PL_Y + «Решение»",
         (not ok2) and all(m in t2 for m in маркеры), t2[:600])

# 4. правило 4 не сломано и идёт первым
def _c4():
    obj = pl_ref[PL_X].ПолучитьОбъект()
    стр = obj.Статьи.Добавить()
    стр.СтатьяДвиженияДенежныхСредств = dds_ref[ДДС_занятая]
    obj.Записать()
ok4, t4 = write_in_transaction([_c4]) if (PL_X and ДДС_занятая) else (True, "данные не подобраны")
проверка("4. Занятая ДДС с оборотами отказана текстом правила 4 (ДДС), не правила 5",
         (not ok4) and ("ДДС не может использоваться" in t4) and ("Статья расходов не может" not in t4), t4[:300])

# 5. shared без оборотов
def _c5_action(p, d):
    def act():
        obj = pl_ref[p].ПолучитьОбъект()
        стр = obj.Статьи.Добавить()
        стр.СтатьяДвиженияДенежныхСредств = dds_ref[d]
        obj.Записать()
    return act
кандидат5 = bool(PL_X and PL_W and СР_без_оборотов and кейс5_действия)
ok5, t5 = (write_in_transaction([_c5_action(p, d) for p, d in кейс5_действия]) if кандидат5
           else (False, "кандидат не найден ни одной из стратегий A/B/C"))
проверка(f"5. Одна СтатьяРасходов без оборотов в двух PL — запись проходит (стратегия {стратегия5 or '—'})",
         кандидат5 and ok5, t5[:300])

# 6. база без следов
after = query("""
ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Строк ИЗ Справочник.А_Статьи_PL.Статьи КАК ТЧ ГДЕ НЕ ТЧ.Ссылка.ПометкаУдаления
""").Получить(0).Строк
проверка("6. База без следов: число строк ТЧ не изменилось после прогона",
         int(after) == int(ТЧ_строк_до), f"до={ТЧ_строк_до}, после={after}")

print(f"\nПРОВЕРОК: {ПРОВЕРОК}, ПРОВАЛОВ: {ПРОВАЛОВ}")
sys.exit(1 if ПРОВАЛОВ else 0)
