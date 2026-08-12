# -*- coding: utf-8 -*-
"""ЦВБ v3.2: переприв язка связей выписка<->списание переводов 1:1 БЕЗ распроведения.

Меняется ТОЛЬКО реквизит А_СписаниеБезналичныхДенежныхСредств (+А_ПодразделениеОтправитель)
у проведённых выписок «Поступление ДС с другого счета». Движения и проведение не трогаются.

Матчинг в группе (Организация, СчетОтправитель, СчетПолучатель, Сумма, Валюта):
хронологический 1:1 — i-я выписка (по дате) <-> i-е списание (по дате).
Протокол: reports\cvb_perevody_links_<mode>.md

Запуск: python fix_cvb_perevody_links.py [--apply]  (без --apply = dry-run)
"""
import argparse
import datetime
import sys
from pathlib import Path

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

parser = argparse.ArgumentParser()
parser.add_argument("--apply", action="store_true", help="боевое применение (иначе dry-run)")
parser.add_argument("--nach", default="2025-12-01")
parser.add_argument("--kon", default="2026-08-31")
args = parser.parse_args()

NACH = datetime.datetime.fromisoformat(args.nach)
KON = datetime.datetime.fromisoformat(args.kon).replace(hour=23, minute=59, second=59)
REPORTS = Path(__file__).resolve().parent / "reports"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ Сп.Ссылка КАК Док, Сп.Дата КАК Дата, Сп.Организация КАК Орг,
	Сп.БанковскийСчет КАК Ист, Сп.БанковскийСчетПолучатель КАК Пол,
	Сп.СуммаДокумента КАК Сумма, Сп.Валюта КАК Вал
ИЗ Документ.СписаниеБезналичныхДенежныхСредств КАК Сп
ГДЕ Сп.Дата МЕЖДУ &Нач И &Кон И Сп.Проведен
	И Сп.ХозяйственнаяОперация = ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПеречислениеДенежныхСредствНаДругойСчет)
УПОРЯДОЧИТЬ ПО Сп.Дата, Сп.Ссылка
"""
q.SetParameter("Нач", NACH)
q.SetParameter("Кон", KON)
r = q.Execute().Выгрузить()
spis = []
for i in range(r.Количество()):
    row = r.Получить(i)
    spis.append({"ref": row.Док, "date": row.Дата,
                 "key": (str(S(row.Орг)), str(S(row.Ист)), str(S(row.Пол)),
                         float(row.Сумма), str(S(row.Вал)))})

q2 = erp.NewObject("Запрос")
q2.Text = """
ВЫБРАТЬ П.Ссылка КАК Док, П.Дата КАК Дата, П.Организация КАК Орг,
	П.БанковскийСчетОтправитель КАК Ист, П.БанковскийСчет КАК Пол,
	П.СуммаДокумента КАК Сумма, П.Валюта КАК Вал,
	П.А_СписаниеБезналичныхДенежныхСредств КАК Тек
ИЗ Документ.ПоступлениеБезналичныхДенежныхСредств КАК П
ГДЕ П.Дата МЕЖДУ &Нач И &Кон И П.Проведен
	И П.ХозяйственнаяОперация = ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПоступлениеДенежныхСредствСДругогоСчета)
УПОРЯДОЧИТЬ ПО П.Дата, П.Ссылка
"""
q2.SetParameter("Нач", NACH)
q2.SetParameter("Кон", KON)
r2 = q2.Execute().Выгрузить()
vyp = []
for i in range(r2.Количество()):
    row = r2.Получить(i)
    vyp.append({"ref": row.Док, "date": row.Дата, "cur": row.Тек,
                "key": (str(S(row.Орг)), str(S(row.Ист)), str(S(row.Пол)),
                        float(row.Сумма), str(S(row.Вал)))})

print(f"Списаний-переводов: {len(spis)}; выписок-приходов: {len(vyp)}", flush=True)

# Группировка и хронологический 1:1 матчинг
from collections import defaultdict
g_sp = defaultdict(list)
for s_item in spis:
    g_sp[s_item["key"]].append(s_item)
g_vp = defaultdict(list)
for v_item in vyp:
    g_vp[v_item["key"]].append(v_item)

plan = []       # (выписка, целевое списание)
bez_spis = []   # выписки без списания
bez_vyp = []    # списания без выписки
for key in sorted(set(list(g_sp) + list(g_vp)), key=str):
    ss = g_sp.get(key, [])
    vv = g_vp.get(key, [])
    for i in range(max(len(ss), len(vv))):
        if i < len(ss) and i < len(vv):
            plan.append((vv[i], ss[i]))
        elif i < len(vv):
            bez_spis.append(vv[i])
        else:
            bez_vyp.append(ss[i])

changes = [(v, s) for (v, s) in plan
           if str(S(v["cur"])) != str(S(s["ref"]))]
print(f"Пар 1:1: {len(plan)}; из них требуют переприв язки: {len(changes)}")
print(f"Выписок без списания: {len(bez_spis)}; списаний без выписки: {len(bez_vyp)}")

lines = [f"# Переприв язка связей переводов ({'APPLY' if args.apply else 'DRY-RUN'})",
         f"Период {NACH:%d.%m.%Y}–{KON:%d.%m.%Y}. Пар: {len(plan)}, переприв язок: {len(changes)}", ""]
applied = 0
errors = 0
for v_item, s_item in changes:
    line = f"- {S(v_item['ref'])}: {S(v_item['cur']) or '(пусто)'} -> {S(s_item['ref'])}"
    if args.apply:
        try:
            obj = v_item["ref"].ПолучитьОбъект()
            obj.А_СписаниеБезналичныхДенежныхСредств = s_item["ref"]
            obj.А_ПодразделениеОтправитель = s_item["ref"].Подразделение
            obj.ОбменДанными.Загрузка = True  # только реквизит: без перепроведения и подписок
            obj.Записать()
            applied += 1
            line += " [OK]"
        except Exception as e:
            msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
            errors += 1
            line += f" [FAIL {str(msg)[:80]}]"
    lines.append(line)

lines.append("")
lines.append("## Выписки БЕЗ списания (в Бух не оформлен перевод — дооформить)")
for v_item in bez_spis:
    lines.append(f"- {S(v_item['ref'])} | {v_item['key'][1][:25]} -> {v_item['key'][2][:25]} | {v_item['key'][3]:,.2f}")
lines.append("")
lines.append("## Списания БЕЗ выписки (входная сторона в ERP не проведена/не создана)")
for s_item in bez_vyp:
    lines.append(f"- {S(s_item['ref'])} | {s_item['key'][1][:25]} -> {s_item['key'][2][:25]} | {s_item['key'][3]:,.2f}")

REPORTS.mkdir(exist_ok=True)
rep = REPORTS / f"cvb_perevody_links_{'apply' if args.apply else 'dry'}.md"
rep.write_text("\n".join(lines), encoding="utf-8")
print(f"Протокол: {rep}")
if args.apply:
    print(f"Применено: {applied}, ошибок: {errors}")
erp = None
print("DONE")
