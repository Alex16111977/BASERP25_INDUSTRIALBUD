# -*- coding: utf-8 -*-
"""Приёмка доработки А_СинхронизироватьДеньги «видимый разлёт» (READ-ONLY, Фазы 1-2).

Кейс-эталон: счёт ОТП, апрель 2026. До доработки Фаза 2 рапортовала
«Документів: 809 (співпало: 809, з діями: 0, потребують рішення бухгалтера: 0)»
при разнице по счёту 17 500 000 — виновника было не видно.

Проверяем:
  1. Пять выписок-переводов, чья пара в Бух проведена маем-июнем, получили статус
     «ПОЗА періодом» и считаются расхождением.
  2. Инвариант замыкания: Разница = ПоясненоДокументами + НеПояснено, причём
     после доработки документы объясняют ВСЮ разницу -> НеПояснено = 0.
  3. Счётчик «співпало» больше не равен общему числу документов.
  4. Регрессия: документы, чья пара проведена В периоде, остались «ОК»/«синхронно».
"""
import datetime
import sys

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
UID_OTP = "ebcb5323-e5fc-11eb-a208-000c299fb278"
OZHIDAEMYE = {
    "DL-00010179": 8000000.0,
    "DL-00010511": 5000000.0,
    "DL-00010422": 3000000.0,
    "DL-00010298": 1000000.0,
    "DL-00010362": 500000.0,
}

FAILS = []


def check(cond, msg):
    print(("  [OK]   " if cond else "  [FAIL] ") + msg, flush=True)
    if not cond:
        FAILS.append(msg)


v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN_ERP)
S = erp.String

plug = erp.Обработки.А_СинхронизироватьДеньги.Создать()
# ⚠️ COM теряет день на «полуночных» датах (грабля проекта): datetime(2026,4,30) приезжает
# в 1С как 29.04.2026, и весь последний день месяца выпадает из сверки. Ставим время явно.
plug.НачалоПериода = datetime.datetime(2026, 4, 1, 0, 0, 1)
plug.ОкончаниеПериода = datetime.datetime(2026, 4, 30, 23, 59, 59)
plug.ФильтрБанковскийСчет = erp.Справочники.БанковскиеСчетаОрганизаций.ПолучитьСсылку(
    erp.NewObject("УникальныйИдентификатор", UID_OTP))
print(f"период в обработке: {S(plug.НачалоПериода)} .. {S(plug.ОкончаниеПериода)}",
      flush=True)

print("=== Фаза 1 ===", flush=True)
print("  " + str(S(plug.СравнитьОстатки())), flush=True)

tr = plug.ТаблицаРасхождений
check(tr.Количество() == 1, f"в таблице расхождений 1 строка (счёт ОТП), факт {tr.Количество()}")
if tr.Количество() == 0:
    print("нечего анализировать", flush=True)
    sys.exit(1)

row = tr.Получить(0)
raznica = float(row.Разница)
print(f"  счёт: {S(row.БанковскийСчет)}", flush=True)
print(f"  нач ЕРП/Бух: {row.НачОстатокЕРП} / {row.НачОстатокБух}", flush=True)
print(f"  приход ЕРП/Бух: {row.ПриходЕРП} / {row.ПриходБух}", flush=True)
print(f"  расход ЕРП/Бух: {row.РасходЕРП} / {row.РасходБух}", flush=True)
print(f"  РАЗНИЦА: {raznica:,.2f}", flush=True)
check(abs(raznica - 17500000.0) < 0.005, "разница по счёту = 17 500 000,00")

print("\n=== Фаза 2 ===", flush=True)
print("  " + str(S(plug.АнализироватьДокументы(0))), flush=True)

td = plug.ТаблицаДокументов
vsego = td.Количество()
sovpalo = 0
s_deystviem = 0
resheniye_buh = 0
poza_periodom = []
ne_poyasneno = []
poyasneno_summa = 0.0

for i in range(vsego):
    d = td.Получить(i)
    status = str(S(d.Статус))
    deystvie = str(S(d.Действие))
    est_rash = bool(d.ЕстьРасхождение)
    predst = str(S(d.ДокументЕРП))
    summa_erp = float(d.СуммаЕРП)
    summa_buh = float(d.СуммаБух)

    if not est_rash:
        sovpalo += 1
    elif deystvie.strip():
        s_deystviem += 1
    else:
        resheniye_buh += 1

    if "ПОЗА періодом" in status:
        poza_periodom.append((predst, summa_erp, status))
    if status.startswith("НЕ ПОЯСНЕНО"):
        ne_poyasneno.append((summa_erp, status))
    if predst.strip():
        poyasneno_summa += summa_erp - summa_buh

print(f"  документов: {vsego} (співпало {sovpalo}, з діями {s_deystviem}, "
      f"рішення бухгалтера {resheniye_buh})", flush=True)

print("\n=== 1. Документы «поза періодом» ===", flush=True)
print(f"  найдено: {len(poza_periodom)}", flush=True)
naydeno_nomera = set()
for predst, summa, status in sorted(poza_periodom, key=lambda x: -x[1]):
    print(f"    {summa:>14,.2f}  {predst[:70]}", flush=True)
    print(f"                    {status[:110]}", flush=True)
    for nom in OZHIDAEMYE:
        if nom in predst:
            naydeno_nomera.add(nom)

for nom, summa in OZHIDAEMYE.items():
    check(nom in naydeno_nomera, f"{nom} ({summa:,.0f}) помечен как «поза періодом»")

check(abs(sum(x[1] for x in poza_periodom) - 17500000.0) < 0.005,
      f"Σ документов «поза періодом» = 17 500 000,00 "
      f"(факт {sum(x[1] for x in poza_periodom):,.2f})")

print("\n=== 2. Инвариант замыкания ===", flush=True)
poyasneno_tch = float(row.ПоясненоДокументами)
ne_poyasneno_tch = float(row.НеПояснено)
print(f"  Разница             = {raznica:,.2f}", flush=True)
print(f"  ПоясненоДокументами = {poyasneno_tch:,.2f}", flush=True)
print(f"  НеПояснено          = {ne_poyasneno_tch:,.2f}", flush=True)
check(abs(poyasneno_tch - poyasneno_summa) < 0.005,
      "реквизит ПоясненоДокументами совпал с пересчётом по строкам")
check(abs(raznica - poyasneno_tch - ne_poyasneno_tch) < 0.005,
      "Разница = ПоясненоДокументами + НеПояснено (инвариант сходится)")
check(abs(ne_poyasneno_tch) < 0.005,
      "НеПояснено = 0 — документы объяснили ВСЮ разницу счёта")
check(len(ne_poyasneno) == 0,
      f"служебных строк «НЕ ПОЯСНЕНО» нет (факт {len(ne_poyasneno)})")
for summa, status in ne_poyasneno:
    print(f"    {summa:>14,.2f}  {status[:110]}", flush=True)

print("\n=== 3. Счётчик «співпало» ===", flush=True)
check(sovpalo < vsego,
      f"«співпало» ({sovpalo}) меньше общего числа документов ({vsego}) — "
      f"разлёт больше не невидим")
check(resheniye_buh >= 5,
      f"«потребують рішення бухгалтера» >= 5 (факт {resheniye_buh})")

print("\n=== 4. Регрессия: документы периода остались ОК ===", flush=True)
ok_v_periode = 0
for i in range(vsego):
    d = td.Получить(i)
    status = str(S(d.Статус))
    if status == "ОК" or status == "Переказ: синхронно" or status.startswith("ОК:"):
        ok_v_periode += 1
check(ok_v_periode >= 700,
      f"документы, чья пара проведена В периоде, остались ОК/синхронно "
      f"(факт {ok_v_periode})")

print(f"\nИТОГ: FAIL {len(FAILS)}", flush=True)
for f in FAILS:
    print(f"  FAIL: {f}", flush=True)

plug = None
erp = None
print("ACCEPTANCE OK" if not FAILS else "ACCEPTANCE FAILED", flush=True)
sys.exit(1 if FAILS else 0)
