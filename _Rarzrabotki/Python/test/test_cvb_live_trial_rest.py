# -*- coding: utf-8 -*-
"""ЖИВОЙ пробный прогон контуров Товары (детально по остаткам), Взаиморасчёты, Касса.

Схема: фазы 1-2 → авто ТОЛЬКО регистрации на обмен (из COM безопасно) → полный цикл
обмена по плану группы → ре-сверка. «Провести в ЕРП» из COM не выполняем (списком).
Товары: разбор каждой расходящейся позиции — историческое (нач. остаток 01.07 уже
расходился) или новое за июль."""
import datetime
import subprocess
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

BASE = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки"
WORKER = r"C:\!OBMIN-OLD-NEW\Обработки\А_ЗапуститьОбмен.epf"
V8EXE = r"C:\Program Files\1cv8\8.3.20.1914\bin\1cv8.exe"
CVB = BASE + r"\А_ЦентрВыравниванияБаз.epf"
D1 = datetime.datetime(2026, 7, 1)
D2 = datetime.datetime(2026, 7, 19)

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

EXCH_ACTS = ("На обмін в BuhBud", "На обмін в Казну", "Переобмін")
POST_ACTS = ("Провести в ЕРП",)


def make(epf):
    p = erp.ВнешниеОбработки.Создать(BASE + "\\" + epf, False)
    p.НачалоПериода = D1
    p.ОкончаниеПериода = D2
    return p


def phase1(p, keyfields, label):
    итог = str(p.СравнитьОстатки())
    rows = []
    total = 0.0
    hist = new = 0
    for i in range(p.ТаблицаРасхождений.Количество()):
        r = p.ТаблицаРасхождений.Получить(i)
        d = float(r.Разница)
        if abs(d) < 0.0005:
            continue
        name = " | ".join(S(getattr(r, f))[:38] for f in keyfields)
        nach_delta = float(r.НачОстатокЕРП) - float(getattr(r, "НачОстаток" + SUF))
        historical = abs(nach_delta) > 0.0005
        if historical:
            hist += 1
        else:
            new += 1
        rows.append((name, d, nach_delta, historical))
        total += abs(d)
    print("%s: %s | ключей: %d (историч. с 01.07: %d, новых за июль: %d) | Σ|Δ| = %.3f"
          % (label, итог, len(rows), hist, new, total))
    return rows, total


def phase23(p, label):
    print("Фаза 2 (%s): %s" % (label, str(p.АнализироватьДокументы(-1))))
    by_action = {}
    n = p.ТаблицаДокументов.Количество()
    for i in range(n):
        a = str(p.ТаблицаДокументов.Получить(i).Действие).strip()
        by_action[a or "(пусто)"] = by_action.get(a or "(пусто)", 0) + 1
    for a, cnt in sorted(by_action.items(), key=lambda x: -x[1]):
        print("   действие %-28r : %d" % (a, cnt))
    reg = 0
    posts = []
    for i in range(n):
        r = p.ТаблицаДокументов.Получить(i)
        a = str(r.Действие).strip()
        if a in EXCH_ACTS:
            r.Синхронизировать = True
            reg += 1
        else:
            r.Синхронизировать = False
            if a in POST_ACTS:
                posts.append(S(r.ДокументЕРП)[:60])
    print("   к регистрации на обмен: %d; отложено «Провести в ЕРП»: %d" % (reg, len(posts)))
    for s in posts[:8]:
        print("      провести: %s" % s)
    if reg:
        print("   Фаза 3: %s" % str(p.ВыполнитьСинхронизацию()))
        ok = sum(1 for i in range(p.ТаблицаДокументов.Количество())
                 if str(p.ТаблицаДокументов.Получить(i).Статус).startswith("Зареєстровано"))
        print("   зарегистрировано фактически: %d" % ok)
    # проблемные без действия
    prob = {}
    for i in range(p.ТаблицаДокументов.Количество()):
        r = p.ТаблицаДокументов.Получить(i)
        if bool(r.ЕстьРасхождение) and not str(r.Действие).strip():
            st = str(r.Статус).strip()[:55]
            prob[st] = prob.get(st, 0) + 1
    for st, cnt in sorted(prob.items(), key=lambda x: -x[1])[:6]:
        print("   без действия: %-55s : %d" % (st, cnt))
    return reg


def chain(base_name, step_label):
    ok_all = True
    for step in (1, 3):
        if step == 3:
            c2 = erp.ВнешниеОбработки.Создать(CVB, False)
            plan = "ОбменУправлениеПредприятиемБухгалтерия20" if base_name == "bas_industrialbud" else KAZNA_PLAN
            t0 = time.time()
            ok = bool(c2.СинхронизироватьПоПлану(plan))
            print("   обмен шаг 2 (ERP, %s): %s, %.0f сек" % (step_label, "OK" if ok else "ОШИБКА", time.time() - t0))
            ok_all = ok_all and ok
        cmd = [V8EXE, "ENTERPRISE", "/S", "localhost\\" + base_name, "/N", "cfo", "/P", "2442",
               "/DisableStartupDialogs", "/DisableStartupMessages", "/Execute", WORKER]
        t0 = time.time()
        rc = subprocess.run(cmd, timeout=900).returncode
        print("   обмен шаг %d (клиент %s): код %d, %.0f сек" % (step, base_name, rc, time.time() - t0))
        ok_all = ok_all and (rc == 0)
    return ok_all


# ---- метаданные плана Казны (без литерала) ----
KAZNA_PLAN = None
mp = erp.Metadata.ExchangePlans
for i in range(mp.Count()):
    nm = str(mp.Get(i).Name)
    if nm.startswith("Казначе"):
        KAZNA_PLAN = nm

print("=" * 78)
print("ГРУППА БУХГАЛТЕРИИ: Товары + Взаиморасчёты, период %s — %s" % (D1.date(), D2.date()))
print("=" * 78)

# ===== ТОВАРЫ =====
SUF = "Бух"
pt = make("СинхронизироватьТоварыТолькоТовары.epf")
rows_t, tot_t = phase1(pt, ("Номенклатура", "Склад"), "ТОВАРЫ ДО")
print("   Расходящиеся позиции (номенклатура | склад | Δкол | Δнач 01.07 | тип):")
for name, d, nd, h in sorted(rows_t, key=lambda x: -abs(x[1])):
    print("   %-78s %10.3f %10.3f  %s" % (name, d, nd, "ИСТОРИЧ." if h else "июль"))
reg_t = phase23(pt, "Товары")
pt = None

# ===== ВЗАИМОРАСЧЁТЫ =====
pv = make("СинхронизироватьВзаиморасчеты.epf")
rows_v, tot_v = phase1(pv, ("Контрагент", "Договор"), "ВЗАИМОРАСЧЁТЫ ДО")
print("   Топ-10 по |Δ|:")
for name, d, nd, h in sorted(rows_v, key=lambda x: -abs(x[1]))[:10]:
    print("   %-78s %12.2f %12.2f  %s" % (name, d, nd, "ИСТОРИЧ." if h else "июль"))
reg_v = phase23(pv, "Взаиморасчёты")
pv = None

# ===== ОБМЕН группы Бух (если были регистрации) =====
if reg_t + reg_v > 0:
    print("-" * 78)
    print("Обмен по плану Бухгалтерии (регистраций: %d)..." % (reg_t + reg_v))
    chain("bas_industrialbud", "план Бух")
else:
    print("Регистраций в группе Бух нет — обмен пропущен.")

# ===== РЕ-СВЕРКА группы Бух =====
pt2 = make("СинхронизироватьТоварыТолькоТовары.epf")
rows_t2, tot_t2 = phase1(pt2, ("Номенклатура", "Склад"), "ТОВАРЫ ПОСЛЕ")
pt2 = None
pv2 = make("СинхронизироватьВзаиморасчеты.epf")
rows_v2, tot_v2 = phase1(pv2, ("Контрагент", "Договор"), "ВЗАИМОРАСЧЁТЫ ПОСЛЕ")
pv2 = None

print("=" * 78)
print("ГРУППА КАЗНЫ: Касса")
print("=" * 78)

# ===== КАССА =====
SUF = "Казна"
pk = make("СинхронизироватьДеньгиКасса.epf")
rows_k, tot_k = phase1(pk, ("Касса",), "КАССА ДО")
for name, d, nd, h in sorted(rows_k, key=lambda x: -abs(x[1])):
    print("   %-50s %14.2f %14.2f  %s" % (name, d, nd, "ИСТОРИЧ." if h else "июль"))
reg_k = phase23(pk, "Касса")
pk = None

if reg_k > 0:
    print("Обмен по плану Казны (регистраций: %d)..." % reg_k)
    chain("kazna", "план Казны")
else:
    print("Регистраций по Кассе нет — обмен пропущен.")

pk2 = make("СинхронизироватьДеньгиКасса.epf")
rows_k2, tot_k2 = phase1(pk2, ("Касса",), "КАССА ПОСЛЕ")
pk2 = None

print("=" * 78)
print("СВОДКА ДО -> ПОСЛЕ:")
print("  Товары:        ключей %3d -> %3d | Σ|Δ| %14.3f -> %14.3f" % (len(rows_t), len(rows_t2), tot_t, tot_t2))
print("  Взаиморасчёты: ключей %3d -> %3d | Σ|Δ| %14.2f -> %14.2f" % (len(rows_v), len(rows_v2), tot_v, tot_v2))
print("  Касса:         ключей %3d -> %3d | Σ|Δ| %14.2f -> %14.2f" % (len(rows_k), len(rows_k2), tot_k, tot_k2))
print("TRIAL REST DONE")
