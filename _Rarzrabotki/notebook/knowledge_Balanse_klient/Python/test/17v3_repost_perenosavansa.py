# -*- coding: utf-8 -*-
"""
СКРИПТ 17v3 — Перепроведение документов с ХозОп=ПереносАванса

Источник: 15v3_docs_to_repost.csv (204 уникальных документа за 13 месяцев)
Паттерн: скрипт 29 (Get-or-Create через Номер+диапазон Даты, Записать(Проведение))
"""
import sys, io, csv, os, re, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from datetime import datetime, timedelta
from _common import connect_erp, ARTIFACTS_DIR

erp = connect_erp()
РежимПроведения = erp.РежимЗаписиДокумента.Проведение

docs = []
with open(os.path.join(ARTIFACTS_DIR, "15v3_docs_to_repost.csv"),
          encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f, delimiter=";"):
        if r["ТипДок"]:
            docs.append((r["ТипДок"], r["ДокИмя"]))

total = len(docs)
print(f"Документов к перепроведению: {total}")
print(f"Старт: {datetime.now().strftime('%H:%M:%S')}\n")

def parse_name(имя):
    m = re.search(r"\s([\w\-]+)\s+от\s+([\d\.]+(?:\s+[\d:]+)?)", имя)
    if m:
        return m.group(1), m.group(2).strip()
    return None, None

def parse_date(s):
    try:
        return datetime.strptime(s, "%d.%m.%Y %H:%M:%S" if " " in s else "%d.%m.%Y")
    except Exception:
        return None

ok = err = 0
errs = []
t0 = time.time()
log_path = os.path.join(ARTIFACTS_DIR, "17v3_progress.log")
log_f = open(log_path, "w", encoding="utf-8")

for i, (тип, имя) in enumerate(docs, 1):
    try:
        ном, дата_стр = parse_name(имя)
        if not ном:
            errs.append((тип, имя, "", "парс имени")); err += 1; continue
        dt = parse_date(дата_стр) if дата_стр else None

        q = erp.NewObject("Запрос")
        q.УстановитьПараметр("Ном", ном)
        if dt is not None:
            q.УстановитьПараметр("Н1", dt - timedelta(days=1))
            q.УстановитьПараметр("Н2", dt + timedelta(days=1))
            q.Текст = f"""ВЫБРАТЬ Ссылка, Дата ИЗ Документ.{тип}
                ГДЕ Номер = &Ном И Дата МЕЖДУ &Н1 И &Н2"""
        else:
            q.Текст = f"ВЫБРАТЬ ПЕРВЫЕ 5 Ссылка, Дата ИЗ Документ.{тип} ГДЕ Номер = &Ном"
        res = q.Выполнить().Выгрузить()
        if res.Количество() == 0:
            errs.append((тип, ном, дата_стр, "не найден")); err += 1; continue

        ref = res.Получить(0).Ссылка
        if res.Количество() > 1 and дата_стр:
            for j in range(res.Количество()):
                rec = res.Получить(j)
                if дата_стр.split()[0] in str(rec.Дата):
                    ref = rec.Ссылка; break

        obj = ref.ПолучитьОбъект()
        if obj is None:
            errs.append((тип, ном, дата_стр, "PolychitObj=None")); err += 1; continue
        t1 = time.time()
        obj.Записать(РежимПроведения)
        dt1 = time.time() - t1
        ok += 1
        print(f"  [{i:3}/{total}] OK  {тип}/{ном} @{дата_стр}  {dt1:.2f}s")
    except Exception as e:
        msg = str(e)[:200]
        if hasattr(e, "excepinfo") and e.excepinfo and e.excepinfo[2]:
            msg = str(e.excepinfo[2])[:200]
        nm = ном if "ном" in dir() else имя
        ds = дата_стр if "дата_стр" in dir() else ""
        errs.append((тип, nm, ds, msg))
        err += 1
        print(f"  [{i:3}/{total}] ERR {тип}/{ном if 'ном' in dir() else имя[:30]}: {msg[:120]}")

    if i % 20 == 0:
        now = time.time(); rate = i / (now - t0)
        log_f.write(f"  [{i}/{total}] OK={ok} ERR={err} скорость={rate:.1f}\n"); log_f.flush()

dt_total = time.time() - t0
final = f"\nИТОГО за {int(dt_total/60)}m{int(dt_total%60)}s: OK={ok}, ERR={err}"
print(final); log_f.write(final + "\n"); log_f.close()

with open(os.path.join(ARTIFACTS_DIR, "17v3_errors.csv"),
          "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["Тип", "Номер", "Дата", "Ошибка"], delimiter=";")
    w.writeheader()
    for t, n, d, m in errs:
        w.writerow({"Тип": str(t), "Номер": str(n), "Дата": str(d), "Ошибка": str(m)[:300]})

sys.exit(0 if err == 0 else 1)
