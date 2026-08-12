# -*- coding: utf-8 -*-
"""Прогон «ВЫРАВНИВАНИЕ ИСТОРИИ»: все 4 контура, период 01.12.2025 — 19.07.2026.

Фаза 2 видит документы всей истории → диагностирует исторические расхождения.
Авто из COM: ТОЛЬКО регистрации на обмен (лимит 500/контур). «Провести в ЕРП» и
класс B (Розпровести/Перепровести) — только списками/счётчиками.
Обмен: одна цепочка на группу плана. Затем ре-сверка всех контуров."""
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


def phase1(p, keyfields, label):
    t0 = time.time()
    итог = str(p.СравнитьОстатки())
    cnt = 0
    total = 0.0
    for i in range(p.ТаблицаРасхождений.Количество()):
        r = p.ТаблицаРасхождений.Получить(i)
        d = float(r.Разница)
        if abs(d) < 0.0005:
            continue
        cnt += 1
        total += abs(d)
    print("%s: %s | ключей: %d | Σ|Δ| = %.3f | %.0f сек"
          % (label, итог, cnt, total, time.time() - t0), flush=True)
    return cnt, total


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
    post_sample = []
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
                if len(post_sample) < 6:
                    post_sample.append(S(r.ДокументЕРП)[:60])
            elif a in B_ACTS:
                n_b += 1
    print("   к регистрации на обмен: %d (лимит %d); «Провести в ЕРП»: %d; класс B: %d"
          % (reg, REG_LIMIT, n_post, n_b), flush=True)
    for s in post_sample:
        print("      провести: %s" % s, flush=True)
    if reg:
        t0 = time.time()
        итог3 = str(p.ВыполнитьСинхронизацию())
        ok = sum(1 for i in range(p.ТаблицаДокументов.Количество())
                 if str(p.ТаблицаДокументов.Получить(i).Статус).startswith("Зареєстровано"))
        errs = {}
        for i in range(p.ТаблицаДокументов.Количество()):
            r = p.ТаблицаДокументов.Получить(i)
            st = str(r.Статус).strip()
            if st.upper().startswith("ПОМИЛКА") and bool(r.Синхронизировать):
                errs[st[:60]] = errs.get(st[:60], 0) + 1
        print("   Фаза 3: %s | зарегистрировано: %d | %.0f сек" % (итог3, ok, time.time() - t0), flush=True)
        for st, cnt in sorted(errs.items(), key=lambda x: -x[1])[:5]:
            print("      ошибка Ф3: %-60s : %d" % (st, cnt), flush=True)
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
print("ВЫРАВНИВАНИЕ ИСТОРИИ: период %s — %s, лимит регистраций %d/контур"
      % (D1.date(), D2.date(), REG_LIMIT), flush=True)
print("=" * 78, flush=True)

results = {}

# ===== ГРУППА БУХГАЛТЕРИИ =====
reg_buh = 0
for epf, keyf, name in (
        ("СинхронизироватьДеньги.epf", ("БанковскийСчет",), "ДЕНЬГИ"),
        ("СинхронизироватьТоварыТолькоТовары.epf", ("Номенклатура", "Склад"), "ТОВАРЫ"),
        ("СинхронизироватьВзаиморасчеты.epf", ("Контрагент", "Договор"), "ВЗАИМОРАСЧЁТЫ")):
    print("-" * 78, flush=True)
    p = make(epf)
    results[name] = {"before": phase1(p, keyf, name + " ДО")}
    reg_buh += phase23(p, name)
    p = None

if reg_buh > 0:
    print("-" * 78, flush=True)
    print("ОБМЕН по плану Бухгалтерии (регистраций: %d)..." % reg_buh, flush=True)
    chain("bas_industrialbud", "ОбменУправлениеПредприятиемБухгалтерия20")
else:
    print("Регистраций в группе Бух нет — обмен пропущен.", flush=True)

for epf, keyf, name in (
        ("СинхронизироватьДеньги.epf", ("БанковскийСчет",), "ДЕНЬГИ"),
        ("СинхронизироватьТоварыТолькоТовары.epf", ("Номенклатура", "Склад"), "ТОВАРЫ"),
        ("СинхронизироватьВзаиморасчеты.epf", ("Контрагент", "Договор"), "ВЗАИМОРАСЧЁТЫ")):
    p = make(epf)
    results[name]["after"] = phase1(p, keyf, name + " ПОСЛЕ")
    p = None

# ===== ГРУППА КАЗНЫ =====
print("=" * 78, flush=True)
p = make("СинхронизироватьДеньгиКасса.epf")
results["КАССА"] = {"before": phase1(p, ("Касса",), "КАССА ДО")}
reg_k = phase23(p, "КАССА")
p = None
if reg_k > 0:
    print("ОБМЕН по плану Казны (регистраций: %d)..." % reg_k, flush=True)
    chain("kazna", KAZNA_PLAN)
else:
    print("Регистраций по Кассе нет — обмен пропущен.", flush=True)
p = make("СинхронизироватьДеньгиКасса.epf")
results["КАССА"]["after"] = phase1(p, ("Касса",), "КАССА ПОСЛЕ")
p = None

print("=" * 78, flush=True)
print("СВОДКА ИСТОРИЯ (01.12.2025—19.07.2026), ДО -> ПОСЛЕ:", flush=True)
for name, rr in results.items():
    b, a = rr["before"], rr["after"]
    print("  %-14s ключей %4d -> %4d | Σ|Δ| %16.3f -> %16.3f"
          % (name + ":", b[0], a[0], b[1], a[1]), flush=True)
print("HISTORY TRIAL DONE", flush=True)
