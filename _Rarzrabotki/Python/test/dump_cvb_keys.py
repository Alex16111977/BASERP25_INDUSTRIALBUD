# -*- coding: utf-8 -*-
"""ЦВБ: дамп ключей расхождений Фазы 1 по всем контурам (READ-ONLY).

Нужен, чтобы снять ТОЧНЫЕ значения `КлючСтроки` для baseline-исключений
(РС А_ИсключенияСверкиБаз): ключ = значения `КонтурСсылка.КлючевыеПоля` через "|"
(контракт движка, `КлючСтрокиРасхождения`).

Ничего не пишет: создаёт плагин, ставит период, вызывает СравнитьОстатки() и печатает
ТаблицаРасхождений (ТЧ обработки, НЕ ТЗ — грабля №1).

Запуск:
  python dump_cvb_keys.py --from 2026-01-01 --to 2026-07-31
  python dump_cvb_keys.py --contour Товары --min-delta 1
"""
import argparse
import datetime
import sys
from pathlib import Path

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
REPORTS = Path(__file__).resolve().parent / "reports"

parser = argparse.ArgumentParser()
parser.add_argument("--from", dest="dt_from", default="2026-01-01")
parser.add_argument("--to", dest="dt_to", default="2026-07-31")
parser.add_argument("--contour", action="append", default=[])
parser.add_argument("--min-delta", type=float, default=0.0005)
parser.add_argument("--out", default="cvb_keys_dump.md")
args = parser.parse_args()

nach = datetime.datetime.strptime(args.dt_from, "%Y-%m-%d")
kon = datetime.datetime.strptime(args.dt_to, "%Y-%m-%d").replace(hour=23, minute=59, second=59)

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN_ERP)
S = erp.String

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
	К.Ссылка КАК Ссылка,
	К.Наименование КАК Наименование,
	К.ИмяОбработки КАК ИмяОбработки,
	К.КлючевыеПоля КАК КлючевыеПоля,
	К.СсылочныеПоля КАК СсылочныеПоля,
	К.ПоляСумм КАК ПоляСумм
ИЗ
	Справочник.А_КонтурыСверкиБаз КАК К
ГДЕ
	НЕ К.ПометкаУдаления
	И НЕ К.Отключен
УПОРЯДОЧИТЬ ПО
	К.Порядок
"""
konturs = q.Execute().Выгрузить()

out = [f"# ЦВБ: дамп ключей Фазы 1 {args.dt_from} .. {args.dt_to}", ""]
print(f"=== Период {nach:%d.%m.%Y} — {kon:%d.%m.%Y} ===", flush=True)

for i in range(konturs.Количество()):
    row = konturs.Получить(i)
    name = str(S(row.Наименование))
    if args.contour and name not in args.contour:
        continue
    plugin_name = str(S(row.ИмяОбработки))
    key_fields = [f.strip() for f in str(S(row.КлючевыеПоля)).split(",") if f.strip()]
    ref_fields = [f.strip() for f in str(S(row.СсылочныеПоля)).split(",") if f.strip()]
    sum_fields = [f.strip() for f in str(S(row.ПоляСумм)).split(",") if f.strip()]

    print(f"\n--- {name} ({plugin_name}) ключ={key_fields} ---", flush=True)
    out += [f"## {name} — плагин `{plugin_name}`",
            f"КлючевыеПоля: `{','.join(key_fields)}`; СсылочныеПоля: `{','.join(ref_fields)}`",
            ""]

    plugin = getattr(erp.Обработки, plugin_name).Создать()

    tch = plugin.Метаданные().ТабличныеЧасти.Найти("ТаблицаРасхождений")
    cols = [str(S(tch.Реквизиты.Получить(j).Имя)) for j in range(tch.Реквизиты.Количество())]
    print(f"    колонки ТЧ: {cols}", flush=True)
    out += [f"Колонки ТаблицаРасхождений: `{', '.join(cols)}`", ""]

    plugin.НачалоПериода = nach
    plugin.ОкончаниеПериода = kon
    try:
        itog = plugin.СравнитьОстатки()
    except Exception as e:
        info = e.excepinfo[2] if getattr(e, "excepinfo", None) else e
        print(f"    ПОМИЛКА Фази 1: {info}", flush=True)
        out += [f"**ПОМИЛКА Фази 1:** {info}", ""]
        continue
    print(f"    Фаза 1: {str(S(itog))[:160]}", flush=True)
    out += [f"Фаза 1: {str(S(itog))[:300]}", ""]

    rows = []
    tab = plugin.ТаблицаРасхождений
    for j in range(tab.Количество()):
        r = tab.Получить(j)
        try:
            delta = float(r.Разница)
        except Exception:
            delta = 0.0
        if abs(delta) < args.min_delta:
            continue
        key = "|".join(str(S(getattr(r, f))) for f in key_fields)
        show_fields = (["Раздел"] if "Раздел" in cols else []) + [f for f in ref_fields
                                                                 if f in cols]
        refs = " / ".join(f"{f}={str(S(getattr(r, f)))[:48]}" for f in show_fields)
        # ПоляСумм контура — логические имена (КоличествоЕРП/СуммаЕРП); в ТЧ плагина
        # реальные колонки называются НачОстаток*/Приход*/Расход*/Остаток* — печатаем ВСЕ,
        # иначе не видно, чья сторона нулевая (это и есть диагноз «ведётся только в ЕРП»).
        num_cols = [c for c in cols
                    if c.startswith(("НачОстаток", "Приход", "Расход", "Остаток"))]
        sums = " ".join(f"{c}={getattr(r, c)}" for c in (num_cols or sum_fields))
        rows.append((abs(delta), delta, key, refs, sums))

    rows.sort(key=lambda x: -x[0])
    print(f"    строк с дельтой: {len(rows)}", flush=True)
    out += [f"Строк с |Δ| >= {args.min_delta}: **{len(rows)}**", "",
            "| Δ | Ключ | Представление | Суммы |", "|---|---|---|---|"]
    for _, delta, key, refs, sums in rows:
        print(f"      Δ={delta:>16,.2f}  {key}  ||  {refs}", flush=True)
        out.append(f"| {delta:,.2f} | `{key}` | {refs} | {sums} |")
    out.append("")

REPORTS.mkdir(exist_ok=True)
rep = REPORTS / args.out
rep.write_text("\n".join(out), encoding="utf-8")
print(f"\nОтчёт: {rep}", flush=True)

erp = None
print("DUMP DONE", flush=True)
