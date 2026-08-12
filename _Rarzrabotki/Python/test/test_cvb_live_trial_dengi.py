# -*- coding: utf-8 -*-
"""ЖИВОЙ пробный прогон выравнивания контура «Деньги безнал» (ERP ↔ bas_industrialbud).

Цикл: Фаза1 (ДО) → Фаза2 → авто-класс ТОЛЬКО регистрации на обмен → полный цикл обмена
(bas_industrialbud → ERP → bas_industrialbud) → Фаза1 повторно (ПОСЛЕ) → сравнение.

«Провести в ЕРП» из COM НЕ выполняем (РегистраторРасчетов создаётся только в UI-сессии) —
эти строки выводим списком. Розпровести/Перепровести — класс B, не трогаем.
"""
import datetime
import subprocess
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

PLUGIN = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьДеньги.epf"
CVB = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\А_ЦентрВыравниванияБаз.epf"
WORKER = r"C:\!OBMIN-OLD-NEW\Обработки\А_ЗапуститьОбмен.epf"
V8EXE = r"C:\Program Files\1cv8\8.3.20.1914\bin\1cv8.exe"
PLAN_BUH = "ОбменУправлениеПредприятиемБухгалтерия20"
D1 = datetime.datetime(2026, 7, 1)
D2 = datetime.datetime(2026, 7, 19)

ACT_EXCHANGE = ("На обмін в BuhBud", "Переобмін")
ACT_POST = ("Провести в ЕРП",)
ACT_B = ("Розпровести в ЕРП", "Перепровести в ЕРП")

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')


def snapshot(label):
    """Фаза 1 свежим объектом плагина -> (число расхождений, Σ|Δ|, [(счёт, дельта)...])."""
    p = erp.ВнешниеОбработки.Создать(PLUGIN, False)
    p.НачалоПериода = D1
    p.ОкончаниеПериода = D2
    итог = str(p.СравнитьОстатки())
    rows = []
    total = 0.0
    for i in range(p.ТаблицаРасхождений.Количество()):
        r = p.ТаблицаРасхождений.Получить(i)
        d = float(r.Разница)
        if abs(d) < 0.005:
            continue
        rows.append((str(r.БанковскийСчет)[:60], d))
        total += abs(d)
    print("%s: %s | расходится счетов: %d | Σ|Δ| = %.2f" % (label, итог, len(rows), total))
    for name, d in sorted(rows, key=lambda x: -abs(x[1])):
        print("   %-60s %15.2f" % (name, d))
    return p, len(rows), total


def run_client(step):
    cmd = [V8EXE, "ENTERPRISE", "/S", r"localhost\bas_industrialbud", "/N", "cfo", "/P", "2442",
           "/DisableStartupDialogs", "/DisableStartupMessages", "/Execute", WORKER]
    t0 = time.time()
    rc = subprocess.run(cmd, timeout=900).returncode
    print("Обмен, шаг %d (клиент bas_industrialbud): код возврата %d, %.0f сек"
          % (step, rc, time.time() - t0))
    return rc == 0


print("=" * 78)
print("ПРОБНЫЙ ПРОГОН: контур «Деньги безнал», период %s — %s" % (D1.date(), D2.date()))
print("=" * 78)

# --- ДО ---
plug, keys_before, total_before = snapshot("ДО  (Фаза 1)")

# --- Фаза 2 на том же объекте ---
print("-" * 78)
итог2 = str(plug.АнализироватьДокументы(-1))
print("Фаза 2:", итог2)

by_action = {}
n = plug.ТаблицаДокументов.Количество()
for i in range(n):
    r = plug.ТаблицаДокументов.Получить(i)
    a = str(r.Действие).strip()
    by_action.setdefault(a or "(пусто)", []).append(i)
for a, idxs in sorted(by_action.items(), key=lambda x: -len(x[1])):
    print("   действие %-28r : %d строк" % (a, len(idxs)))

# --- флаги: авто = ТОЛЬКО регистрации на обмен ---
reg, post_skipped, b_skipped = 0, [], []
for i in range(n):
    r = plug.ТаблицаДокументов.Получить(i)
    a = str(r.Действие).strip()
    if a in ACT_EXCHANGE:
        r.Синхронизировать = True
        reg += 1
    else:
        r.Синхронизировать = False
        if a in ACT_POST:
            post_skipped.append("%s | %s" % (str(r.ДокументЕРП)[:70], str(r.Статус)[:50]))
        elif a in ACT_B:
            b_skipped.append("%s | %s | %s" % (a, str(r.ДокументЕРП)[:60], str(r.Статус)[:40]))

print("-" * 78)
print("К авто-выполнению (регистрации на обмен): %d строк" % reg)
print("Отложено «Провести в ЕРП» (только UI-сессия): %d" % len(post_skipped))
for s in post_skipped[:10]:
    print("   |", s)
print("Отложено класс B (подтверждение): %d" % len(b_skipped))
for s in b_skipped[:10]:
    print("   |", s)

if reg > 0:
    итог3 = str(plug.ВыполнитьСинхронизацию())
    print("Фаза 3:", итог3)
    ok_reg = sum(1 for i in range(plug.ТаблицаДокументов.Количество())
                 if str(plug.ТаблицаДокументов.Получить(i).Статус).startswith("Зареєстровано"))
    print("Фактически зарегистрировано на обмен: %d" % ok_reg)
else:
    print("Регистраций нет — обмен всё равно прогоняем (вдруг хвост с той стороны).")

plug = None  # освободить COM плагина до обмена

# --- ПОЛНЫЙ ЦИКЛ ОБМЕНА: bas_industrialbud -> ERP -> bas_industrialbud ---
print("-" * 78)
ok1 = run_client(1)
cvb = erp.ВнешниеОбработки.Создать(CVB, False)
t0 = time.time()
ok2 = bool(cvb.СинхронизироватьПоПлану(PLAN_BUH))
print("Обмен, шаг 2 (ERP в COM-сессии, план Бухгалтерии): %s, %.0f сек"
      % ("OK" if ok2 else "ОШИБКА", time.time() - t0))
ok3 = run_client(3)
print("Цепочка обмена: %s" % ("OK" if (ok1 and ok2 and ok3) else "ЕСТЬ СБОИ"))

# --- ПОСЛЕ ---
print("-" * 78)
plug2, keys_after, total_after = snapshot("ПОСЛЕ (ре-сверка)")

print("=" * 78)
print("ИТОГ: счетов с расхождением %d -> %d | Σ|Δ| %.2f -> %.2f (изменение %.2f)"
      % (keys_before, keys_after, total_before, total_after, total_before - total_after))
print("TRIAL DONE")
