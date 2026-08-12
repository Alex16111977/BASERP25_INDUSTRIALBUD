# -*- coding: utf-8 -*-
"""ЦВБ v3.2: запуск штатной санации связей переводов (плагин А_СинхронизироватьДеньги).

python run_cvb_sanacia_perevodov.py            -> только отчёт
python run_cvb_sanacia_perevodov.py --apply    -> боевое применение (штатная процедура плагина)
"""
import argparse
import datetime
import sys
from pathlib import Path

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

parser = argparse.ArgumentParser()
parser.add_argument("--apply", action="store_true")
parser.add_argument("--nach", default="2025-12-01")
parser.add_argument("--kon", default="2026-08-31")
args = parser.parse_args()

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

plug = erp.DataProcessors.А_СинхронизироватьДеньги.Create()
plug.НачалоПериода = datetime.datetime.fromisoformat(args.nach)
plug.ОкончаниеПериода = datetime.datetime.fromisoformat(args.kon)

protokol = plug.ВыровнятьСвязиПереводов(not args.apply)
lines = [f"# Санация связей переводов ({'APPLY' if args.apply else 'ОТЧЁТ'}) "
         f"{args.nach}..{args.kon}", ""]
# Литералы протокола (ObjectModule А_СинхронизироватьДеньги, ВыровнятьСвязиПереводов):
#   "Множинне списання …"       — списание с >1 проведённой выпиской
#   "[звіт] Прив'язка: … -> …"  — привязка (в отчётном режиме с префиксом [звіт])
#   "Виписок без прив'язки: N"  — сводная строка шага 3
#   "БЕЗ ПАРИ: виписка …"       — свободного списания группы нет (реальный разрыв)
#   "РОЗРИВ: списання без проведеної виписки — …" — итоговый список дыр; в ОТЧЁТНОМ
#       режиме это ДО-состояние (привязки ещё не записаны), в apply — ПОСЛЕ.
n_multi = n_link = n_bezpary = n_razryv = n_err = 0
n_svobodnyh = "?"
for i in range(protokol.Количество()):
    line = str(S(protokol.Получить(i)))
    lines.append("- " + line)
    if line.startswith("Множинне"):
        n_multi += 1
    elif "Прив'язка:" in line:
        n_link += 1
    elif line.startswith("Виписок без прив'язки:"):
        n_svobodnyh = line.split(":")[-1].strip()
    elif line.startswith("БЕЗ ПАРИ"):
        n_bezpary += 1
    elif line.startswith("РОЗРИВ"):
        n_razryv += 1
    elif line.startswith("ПОМИЛКА"):
        n_err += 1

rep = Path(__file__).resolve().parent / "reports" / (
    f"cvb_sanacia_perevodov_{'apply' if args.apply else 'report'}.md")
rep.parent.mkdir(exist_ok=True)
rep.write_text("\n".join(lines), encoding="utf-8")
svod = (f"множественных списаний: {n_multi}; привязок: {n_link}; "
        f"свободных выписок (шаг 3): {n_svobodnyh}; выписок без пары: {n_bezpary}; "
        f"разрывов (списание без проведённой выписки): {n_razryv}; ошибок записи: {n_err}")
lines.insert(1, "")
lines.insert(2, svod)
rep.write_text("\n".join(lines), encoding="utf-8")
print(svod)
print(f"Протокол: {rep}")
plug = None
erp = None
print("DONE")
