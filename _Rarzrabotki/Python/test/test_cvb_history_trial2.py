# -*- coding: utf-8 -*-
"""Прогон «ВЫРАВНИВАНИЕ ИСТОРИИ» v2 — С ОПТИМИЗАЦИЕЙ Фазы 2 (спека §1.8.2):
после Фазы 1 отметка остаётся только у строк с РЕАЛЬНОЙ дельтой остатка и только
у сводных складов/контрагентов, где такие строки есть. Деньги уже продиагностированы
в v1 (4 регистрации сделаны) — здесь только обмен и ре-сверка.
Период 01.12.2025 — 19.07.2026. Авто из COM: только регистрации на обмен (лимит 500)."""
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
D1 = datetime.datetime(2025, 12, 1)
D2 = datetime.datetime(2026, 7, 19)
REG_LIMIT = 500
THRESH = 0.0005

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

EXCH_ACTS = ("На обмін в BuhBud", "На обмін в Казну", "Переобмін")
POST_ACTS = ("Провести в ЕРП",)
B_ACTS = ("Розпровести в ЕРП", "Перепровести в ЕРП")

KAZNA_PLAN = None
mp = erp.Metadata.ExchangePlans
for i in range(mp.Count()):
    nm = str(mp.Get(i).Name)
    if nm.startswith("Казначе"):
        KAZNA_PLAN = nm


def make(epf):
    p = erp.ВнешниеОбработки.Создать(BASE + "\\" + epf, False)
    p.НачалоПериода = D1
    p.ОкончаниеПериода = D2
    return p


def phase1(p, label):
    t0 = time.time()
    итог = str(p.СравнитьОстатки())
    cnt = 0
    total = 0.0
    for i in range(p.ТаблицаРасхождений.Количество()):
        r = p.ТаблицаРасхождений.Получить(i)
        d = float(r.Разница)
        if abs(d) < THRESH:
            continue
        cnt += 1
        total += abs(d)
    print("%s: %s | реальных ключей: %d | Σ|Δ| = %.3f | %.0f сек"
          % (label, итог, cnt, total, time.time() - t0), flush=True)
    return cnt, total


def optimize_flags(p, svod_name, svod_key):
    """§1.8.2: отметка только реальным строкам и их сводам. Возвращает (реальных строк, отмеч. сводов)."""
    real_keys = set()
    for i in range(p.ТаблицаРасхождений.Количество()):
        r = p.ТаблицаРасхождений.Получить(i)
        if abs(float(r.Разница)) < THRESH:
            r.Синхронизировать = False
        else:
            r.Синхронизировать = True
            real_keys.add(str(getattr(r, svod_key)))
    svod = getattr(p, svod_name)
    marked = 0
    for i in range(svod.Количество()):
        row = svod.Получить(i)
        on = str(getattr(row, svod_key)) in real_keys
        row.Синхронизировать = on
        if on:
            marked += 1
    print("   оптимизация: реальных строк %d, отмечено сводов %d из %d"
          % (len(real_keys), marked, svod.Количество()), flush=True)
    return marked


def phase23(p, label):
    t0 = time.time()
    итог2 = str(p.АнализироватьДокументы(-1))
    print("Фаза 2 (%s): %s | %.0f сек" % (label, итог2, time.time() - t0), flush=True)
    by_action = {}
    n = p.ТаблицаДокументов.Количество()
    for i in range(n):
        a = str(p.ТаблицаДокументов.Получить(i).Действие).strip()
        by_action[a or "(пусто)"] = by_action.get(a or "(пусто)", 0) + 1
    for a, cnt in sorted(by_action.items(), key=lambda x: -x[1]):
        print("   действие %-28r : %d" % (a, cnt), flush=True)
    reg = 0
    n_post = n_b = 0
    for i in range(n):
        r = p.ТаблицаДокументов.Получить(i)
        a = str(r.Действие).strip()
        if a in EXCH_ACTS and reg < REG_LIMIT:
            r.Синхронизировать = True
            reg += 1
        else:
            r.Синхронизировать = False
            if a in POST_ACTS:
                n_post += 1
            elif a in B_ACTS:
                n_b += 1
    print("   к регистрации: %d; «Провести в ЕРП»: %d; класс B: %d" % (reg, n_post, n_b), flush=True)
    if reg:
        t0 = time.time()
        итог3 = str(p.ВыполнитьСинхронизацию())
        ok = sum(1 for i in range(p.ТаблицаДокументов.Количество())
                 if str(p.ТаблицаДокументов.Получить(i).Статус).startswith("Зареєстровано"))
        print("   Фаза 3: %s | зарегистрировано: %d | %.0f сек" % (итог3, ok, time.time() - t0), flush=True)
    return reg


def chain(base_name, plan):
    ok_all = True
    for step in (1, 3):
        if step == 3:
            c2 = erp.ВнешниеОбработки.Создать(CVB, False)
            t0 = time.time()
            ok = bool(c2.СинхронизироватьПоПлану(plan))
            print("   обмен шаг 2 (ERP): %s, %.0f сек" % ("OK" if ok else "ОШИБКА", time.time() - t0), flush=True)
            ok_all = ok_all and ok
        cmd = [V8EXE, "ENTERPRISE", "/S", "localhost\\" + base_name, "/N", "cfo", "/P", "2442",
               "/DisableStartupDialogs", "/DisableStartupMessages", "/Execute", WORKER]
        t0 = time.time()
        rc = subprocess.run(cmd, timeout=1800).returncode
        print("   обмен шаг %d (клиент %s): код %d, %.0f сек" % (step, base_name, rc, time.time() - t0), flush=True)
        ok_all = ok_all and (rc == 0)
    return ok_all


print("=" * 78, flush=True)
print("ИСТОРИЯ v2 (оптимизация §1.8.2): %s — %s" % (D1.date(), D2.date()), flush=True)
print("=" * 78, flush=True)

res = {}

# ТОВАРЫ — с оптимизацией сводов
print("-" * 78, flush=True)
p = make("СинхронизироватьТоварыТолькоТовары.epf")
res["ТОВАРЫ"] = {"before": phase1(p, "ТОВАРЫ ДО")}
optimize_flags(p, "ТаблицаСкладов", "СкладКлюч")
reg_t = phase23(p, "ТОВАРЫ")
p = None

# ВЗАИМОРАСЧЁТЫ — с оптимизацией сводов
print("-" * 78, flush=True)
p = make("СинхронизироватьВзаиморасчеты.epf")
res["ВЗАИМОРАСЧЁТЫ"] = {"before": phase1(p, "ВЗАИМОРАСЧЁТЫ ДО")}
optimize_flags(p, "ТаблицаКонтрагентов", "КонтрагентКлюч")
reg_v = phase23(p, "ВЗАИМОРАСЧЁТЫ")
p = None

# ОБМЕН группы Бух (несёт и 4 денежные регистрации из v1)
print("-" * 78, flush=True)
print("ОБМЕН по плану Бухгалтерии (регистраций этого прогона: %d + 4 денежные из v1)..."
      % (reg_t + reg_v), flush=True)
chain("bas_industrialbud", "ОбменУправлениеПредприятиемБухгалтерия20")

# Ре-сверка группы Бух (+деньги)
for epf, name in (("СинхронизироватьДеньги.epf", "ДЕНЬГИ"),
                  ("СинхронизироватьТоварыТолькоТовары.epf", "ТОВАРЫ"),
                  ("СинхронизироватьВзаиморасчеты.epf", "ВЗАИМОРАСЧЁТЫ")):
    p = make(epf)
    res.setdefault(name, {})["after"] = phase1(p, name + " ПОСЛЕ")
    p = None

# КАССА
print("=" * 78, flush=True)
p = make("СинхронизироватьДеньгиКасса.epf")
res["КАССА"] = {"before": phase1(p, "КАССА ДО")}
reg_k = phase23(p, "КАССА")
p = None
if reg_k > 0:
    print("ОБМЕН по плану Казны (регистраций: %d)..." % reg_k, flush=True)
    chain("kazna", KAZNA_PLAN)
else:
    print("Регистраций по Кассе нет — обмен пропущен.", flush=True)
p = make("СинхронизироватьДеньгиКасса.epf")
res["КАССА"]["after"] = phase1(p, "КАССА ПОСЛЕ")
p = None

print("=" * 78, flush=True)
print("СВОДКА ИСТОРИЯ v2, ДО -> ПОСЛЕ:", flush=True)
for name, rr in res.items():
    b = rr.get("before")
    a = rr.get("after")
    if b and a:
        print("  %-14s ключей %4d -> %4d | Σ|Δ| %16.3f -> %16.3f"
              % (name + ":", b[0], a[0], b[1], a[1]), flush=True)
    elif a:
        print("  %-14s ПОСЛЕ: ключей %4d | Σ|Δ| %16.3f (ДО — см. v1)" % (name + ":", a[0], a[1]), flush=True)
print("HISTORY TRIAL V2 DONE", flush=True)
