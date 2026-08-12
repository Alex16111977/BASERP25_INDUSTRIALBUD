# -*- coding: utf-8 -*-
"""ЦВБ v3: оркестратор помесячного выравнивания (январь..июль 2026).

Конвейер месяца (спека 2026-08-08 §3):
  1. Полный обмен всех баз (worker-клиенты 1cv8 + обмен ERP через COM).
  2. Сеанс 4 контуров за месяц: РежимИстория=Истина («Провести» = класс B, авто
     выполняет ТОЛЬКО регистрации обмена + санацию — безопасно из COM).
  3. Если были регистрации — полный обмен и ре-сверка (до --max-iter, анти-зацикливание).
  4. Отчёт месяца: _Rarzrabotki\\Python\\test\\reports\\cvb_<YYYY-MM>.md.

Примеры:
  python run_cvb_month.py --month 2026-01 --report-only
  python run_cvb_month.py --month 2026-01
  python run_cvb_month.py --month 2026-03 --contour "Товары"
"""
import argparse
import calendar
import datetime
import subprocess
import sys
import time
from pathlib import Path

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

V8_EXE = r"C:\Program Files\1cv8\8.3.20.1914\bin\1cv8.exe"
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
REPORTS = Path(__file__).resolve().parent / "reports"
ALL_CONTOURS = ["Деньги безнал", "Товары", "Взаиморасчёты", "Касса"]

parser = argparse.ArgumentParser()
parser.add_argument("--month", required=True, help="YYYY-MM")
parser.add_argument("--contour", action="append", default=[],
                    help="Только указанные контуры (повторяемый)")
parser.add_argument("--report-only", action="store_true",
                    help="Фазы 1-2 + классификация, БЕЗ обмена и Фазы 3")
parser.add_argument("--skip-exchange", action="store_true",
                    help="Не гонять полный обмен (для повторной сверки)")
parser.add_argument("--max-iter", type=int, default=3)
parser.add_argument("--limit-reg", type=int, default=500)
parser.add_argument("--limit-podtv", type=int, default=200)
args = parser.parse_args()

year, month = map(int, args.month.split("-"))

# ⚠️ COM-конвертер сдвигает КАЖДУЮ дату на −3 часа (часовой пояс клиента UTC+3 → сервер),
# а реквизит типа Дата у плагина вдобавок отсекает время. Из-за этого datetime(2026,4,1)
# приезжает в 1С как 31.03.2026, и месяц сверялся по окну, сдвинутому на сутки назад
# (проверено probe: 01.04 00:00 -> 31.03; 01.04 03:00 -> 01.04 00:00).
# Компенсируем сдвиг: +3 часа — тогда в 1С попадает ровно нужная дата 00:00:00.
COM_SDVIG = datetime.timedelta(hours=3)
nach_pokaz = datetime.datetime(year, month, 1)
kon_pokaz = datetime.datetime(year, month, calendar.monthrange(year, month)[1])
nach = nach_pokaz + COM_SDVIG
kon = kon_pokaz + COM_SDVIG
contours = args.contour if args.contour else ALL_CONTOURS

print(f"=== ЦВБ месяц {args.month}: {nach_pokaz:%d.%m.%Y} — {kon_pokaz:%d.%m.%Y} ===",
      flush=True)
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN_ERP)
S = erp.String
engine = erp.А_ВыравниваниеБазСервер


def plan_names():
    names = engine.ВсеИменаПлановОбмена()
    return [str(S(names.Получить(i))) for i in range(names.Количество())]


def worker_step(settings, step_no):
    cmd = [V8_EXE, "ENTERPRISE",
           "/S", f"{settings['Сервер']}\\{settings['ИмяБазы']}",
           "/N", settings["Пользователь"], "/P", settings["Пароль"],
           "/DisableStartupDialogs", "/DisableStartupMessages",
           "/Execute", settings["ПутьКОбработке"]]
    t0 = time.time()
    rc = subprocess.run(cmd, capture_output=True, timeout=900).returncode
    print(f"    шаг {step_no} ({settings['ИмяБазы']}): rc={rc}, {time.time()-t0:.0f} c", flush=True)
    return rc == 0


def full_exchange(log):
    ok_all = True
    for plan in plan_names():
        st = engine.НастройкиОбменаПоИмениПлана(plan)
        settings = {k: str(S(getattr(st, k))) for k in
                    ("Сервер", "ИмяБазы", "Пользователь", "Пароль", "ПутьКОбработке")}
        if not settings["ПутьКОбработке"]:
            log.append(f"  обмен {plan}: НЕТ пути worker'а — пропуск")
            ok_all = False
            continue
        print(f"  обмен по плану {plan} ({settings['ИмяБазы']}) ...", flush=True)
        ok = worker_step(settings, 1)
        msg = engine.СинхронизироватьПоПлану(plan)
        print(f"    шаг 2 (ERP): {str(S(msg))[:120]}", flush=True)
        ok = worker_step(settings, 3) and ok
        log.append(f"  обмен {plan}: {'OK' if ok else 'СБОИ'}")
        ok_all = ok_all and ok
    return ok_all


def kontur_ref(name):
    ref = erp.Справочники.А_КонтурыСверкиБаз.НайтиПоНаименованию(name, True)
    assert str(S(ref)).strip(), f"контур не найден: {name}"
    return ref


def run_kontur(name, avto, iteration, session_dt):
    t0 = time.time()
    res = engine.ОбработатьКонтурПоЗначениям(
        kontur_ref(name), nach, kon, avto, True,
        args.limit_reg, args.limit_podtv, session_dt, iteration)
    m = res.Метрика
    out = {
        "контур": name,
        "ошибка": bool(m.ЕстьОшибка),
        "расходится": int(m.КлючейРасходится),
        "объяснено": int(m.КлючейОбъяснено),
        "дельта": float(m.СуммаДельта),
        "классA": int(m.КлассА),
        "автоисправлено": int(m.Автоисправлено),
        "на_обмен": int(m.НаОбмен),
        "подтв": int(m.КПодтверждению),
        "подтв_сверх": int(m.ПодтвержденийСверхЛимита),
        "разбор": int(m.РучнойРазбор),
        "битых": int(m.БитыхСсылок),
        "регистрации": bool(m.ЕстьРегистрации),
        "сек": round(time.time() - t0),
        "сообщения": [str(S(m.Сообщения.Получить(i)))
                      for i in range(m.Сообщения.Количество())],
        "разбор_строки": [],
        "подтв_строки": [],
    }
    for i in range(res.СтрокиРазбора.Количество()):
        st = res.СтрокиРазбора.Получить(i)
        out["разбор_строки"].append(
            f"{S(st.ТипДокумента)} | {S(st.ДокументПредставление)} | {S(st.Статус)} | "
            f"{S(st.Диагноз)} | Δ {st.Дельта}")
    for i in range(min(res.СтрокиПодтверждения.Количество(), 40)):
        st = res.СтрокиПодтверждения.Получить(i)
        out["подтв_строки"].append(
            f"{S(st.Действие)} | {S(st.ТипДокумента)} | {S(st.ДокументПредставление)} | "
            f"{S(st.Статус)} | Δ {st.Дельта}")
    print(f"  [{name}] расх={out['расходится']} (объясн {out['объяснено']}) "
          f"Δ={out['дельта']:.2f} A={out['классA']} обмен={out['на_обмен']} "
          f"B={out['подтв']}+{out['подтв_сверх']} C={out['разбор']} "
          f"рег={out['регистрации']} за {out['сек']} c", flush=True)
    return out


session_dt = datetime.datetime.now().replace(microsecond=0) + COM_SDVIG
avto = not args.report_only
log_lines = []
results_by_iter = []

if avto and not args.skip_exchange:
    print("--- Полный обмен ДО сверки ---", flush=True)
    full_exchange(log_lines)

prev_metric = None
for iteration in range(1, (args.max_iter if avto else 1) + 1):
    print(f"--- Итерация {iteration} ---", flush=True)
    itog = [run_kontur(name, avto, iteration, session_dt) for name in contours]
    results_by_iter.append(itog)
    total_rash = sum(r["расходится"] for r in itog)
    were_reg = any(r["регистрации"] for r in itog)
    errors = any(r["ошибка"] for r in itog)
    if errors:
        log_lines.append(f"итерация {iteration}: ОШИБКИ фаз — стоп")
        break
    if total_rash == 0:
        log_lines.append(f"итерация {iteration}: расхождений нет — ЗЕЛЁНЫЙ")
        break
    if not avto:
        break
    if prev_metric is not None and total_rash >= prev_metric:
        log_lines.append(f"итерация {iteration}: метрика не улучшается "
                         f"({prev_metric} -> {total_rash}) — стоп")
        break
    prev_metric = total_rash
    if were_reg and not args.skip_exchange:
        print("--- Обмен после регистраций ---", flush=True)
        full_exchange(log_lines)
    else:
        log_lines.append(f"итерация {iteration}: регистраций нет — ре-сверка не нужна")
        break

# --- Отчёт месяца ---
REPORTS.mkdir(exist_ok=True)
rep = REPORTS / f"cvb_{args.month}.md"
lines = [f"# ЦВБ: {args.month} ({'отчёт' if args.report_only else 'выравнивание'})",
         f"Дата сеанса: {session_dt}", ""]
for it_no, itog in enumerate(results_by_iter, 1):
    lines.append(f"## Итерация {it_no}")
    lines.append("| Контур | Расх. | Объясн. | Σ\\|Δ\\| | A | На обмен | B | C | сек |")
    lines.append("|---|---|---|---|---|---|---|---|---|")
    for r in itog:
        lines.append(f"| {r['контур']} | {r['расходится']} | {r['объяснено']} | "
                     f"{r['дельта']:.2f} | {r['классA']} | {r['на_обмен']} | "
                     f"{r['подтв']}+{r['подтв_сверх']} | {r['разбор']} | {r['сек']} |")
    lines.append("")
    for r in itog:
        if r["сообщения"]:
            lines.append(f"### {r['контур']}: сообщения")
            lines.extend(f"- {s}" for s in r["сообщения"])
            lines.append("")
        if r["разбор_строки"]:
            lines.append(f"### {r['контур']}: ручной разбор (класс C)")
            lines.extend(f"- {s}" for s in r["разбор_строки"])
            lines.append("")
        if r["подтв_строки"]:
            lines.append(f"### {r['контур']}: к подтверждению (класс B, топ)")
            lines.extend(f"- {s}" for s in r["подтв_строки"])
            lines.append("")
if log_lines:
    lines.append("## Лог конвейера")
    lines.extend(f"- {s}" for s in log_lines)
rep.write_text("\n".join(lines), encoding="utf-8")
print(f"\nОтчёт: {rep}", flush=True)

erp = None
print("MONTH DONE", flush=True)
