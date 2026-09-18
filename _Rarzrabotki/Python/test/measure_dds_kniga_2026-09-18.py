# -*- coding: utf-8 -*-
"""Замер книги экономиста «ДДС Производство Вересень - Грудень 2026 Прогноз.xlsx».

Ничего не пишет в базу. Читает книгу openpyxl (data_only=True), по каждому листу:
  - видимость, размеры;
  - месячные блоки ПО ДАТАМ строки 3 (не по фиксированным колонкам);
  - шапку строки 4 под каждым блоком;
  - строки данных с уровнем группировки (outline_level), жирностью, отступом;
  - суммы Безнал/Нал по уровням 0/1/2 и по месяцам;
  - знак (есть ли отрицательные), контроль Итого = Безнал + Нал;
  - комментарии блоков, хвост листа после последнего датового блока.
Полный дамп — в JSON рядом (scratchpad), на экран — сводка.
"""
import sys, json, datetime, os
sys.stdout.reconfigure(encoding='utf-8')
from openpyxl import load_workbook

PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Бюджет прогноз\Отчеты высланные\ДДС Производство Вересень - Грудень 2026 Прогноз.xlsx"
OUT_DIR = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
OUT_JSON = os.path.join(OUT_DIR, "dds_kniga_dump.json")

RAW_PREFIX = "Date"   # сырые движения — только размеры, без построчного разбора


def norm(s):
    if s is None:
        return ""
    s = str(s).replace("\xa0", " ").replace("\t", " ").replace("\n", " ").replace("\r", " ")
    while "  " in s:
        s = s.replace("  ", " ")
    return s.strip().lower()


def num(v):
    if v is None or isinstance(v, (bool, datetime.datetime, datetime.date)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(t)
    except Exception:
        return 0.0


def is_date(v):
    return isinstance(v, (datetime.datetime, datetime.date))


def fmt(x):
    return f"{x:,.2f}".replace(",", " ")


t0 = datetime.datetime.now()
wb = load_workbook(PATH, data_only=True)
print(f"Книга открыта за {(datetime.datetime.now() - t0).seconds} с; листов {len(wb.sheetnames)}")

dump = {"file": PATH, "sheets": []}
hidden_cnt = 0

for ws in wb.worksheets:
    state = ws.sheet_state
    if state != "visible":
        hidden_cnt += 1
    info = {
        "name": ws.title, "state": state, "max_row": ws.max_row, "max_col": ws.max_column,
        "top_rows": {}, "blocks": [], "header4": {}, "rows": [], "end_of_month_part_col": None,
        "tail_cols": [], "dates_elsewhere": [],
    }
    print("\n" + "=" * 100)
    print(f"ЛИСТ «{ws.title}»  state={state}  строк={ws.max_row}  колонок={ws.max_column}")

    if ws.title.startswith(RAW_PREFIX):
        # сырые движения: только шапка
        hdr = [(c, ws.cell(1, c).value) for c in range(1, min(ws.max_column, 30) + 1) if ws.cell(1, c).value is not None]
        info["top_rows"]["1"] = [[c, str(v)[:40]] for c, v in hdr]
        print(f"  сырые движения, шапка строки 1: {[str(v)[:25] for _, v in hdr][:12]}")
        dump["sheets"].append(info)
        continue

    # верхние ряды 1..6 — всё непустое
    for r in range(1, 7):
        cells = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip() != "":
                cells.append([c, v.isoformat() if is_date(v) else str(v)[:60]])
        info["top_rows"][str(r)] = cells
        print(f"  ряд {r}: {cells[:14]}{' ...' if len(cells) > 14 else ''}")

    # даты в рядах 1..6 (где угодно) — чтобы поймать другую разметку
    for r in range(1, 7):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if is_date(v):
                info["dates_elsewhere"].append([r, c, v.isoformat()])

    # месячные блоки — по датам строки 3
    blocks = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(3, c).value
        if is_date(v):
            month = datetime.date(v.year, v.month, 1)
            blocks.append({"col": c, "date": v.isoformat(), "month": month.isoformat()})
    info["blocks"] = blocks
    if blocks:
        last = blocks[-1]["col"]
        # конец месячной части: первая колонка после последнего блока, где в строке 3 не дата
        end_col = last + 4
        info["end_of_month_part_col"] = end_col
        tail = []
        for c in range(end_col, ws.max_column + 1):
            vals = []
            for r in range(1, 7):
                v = ws.cell(r, c).value
                if v is not None and str(v).strip() != "":
                    vals.append(f"r{r}={str(v)[:25]}")
            if vals:
                tail.append([c, vals])
        info["tail_cols"] = tail
    print(f"  блоки по датам строки 3: {[(b['col'], b['month']) for b in blocks]}")
    if info["tail_cols"]:
        print(f"  ХВОСТ после месячной части (с колонки {info['end_of_month_part_col']}): {info['tail_cols'][:10]}")

    # шапка строки 4 под блоками
    for b in blocks:
        c = b["col"]
        h = [str(ws.cell(4, c + k).value or "")[:16] for k in range(4)]
        info["header4"][str(c)] = h
    print(f"  шапка ряда 4 по блокам: {info['header4']}")
    if not blocks:
        dump["sheets"].append(info)
        continue

    # строки данных
    by_level = {}   # (level) -> {month: [bn, nal, cnt_nonzero]}
    negatives = []
    itogo_mismatch = []
    names_lvl2 = []
    last_named_row = 0
    for r in range(6, ws.max_row + 1):
        rd = ws.row_dimensions[r]
        level = rd.outline_level or 0
        hidden_row = bool(rd.hidden)
        name = None
        name_col = None
        for c in (1, 2, 3):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip() != "":
                name, name_col = str(v).strip(), c
                break
        a = ws.cell(r, 1)
        bold = bool(a.font.bold) if a.font is not None else False
        indent = a.alignment.indent if a.alignment is not None else 0
        vals = {}
        any_val = False
        for b in blocks:
            c = b["col"]
            bn = num(ws.cell(r, c).value)
            nal = num(ws.cell(r, c + 1).value)
            itogo = num(ws.cell(r, c + 2).value)
            comment = ws.cell(r, c + 3).value
            comment = str(comment).strip()[:80] if comment is not None and str(comment).strip() != "" else None
            raw_bn = ws.cell(r, c).value
            raw_nal = ws.cell(r, c + 1).value
            if bn != 0 or nal != 0 or comment:
                any_val = True
            vals[b["month"]] = {"bn": bn, "nal": nal, "itogo": itogo, "comment": comment,
                                 "raw_bn_type": type(raw_bn).__name__ if raw_bn is not None else None,
                                 "raw_nal_type": type(raw_nal).__name__ if raw_nal is not None else None}
            if bn < 0 or nal < 0:
                negatives.append([r, name, b["month"], bn, nal])
            if abs((bn + nal) - itogo) > 0.011 and (bn != 0 or nal != 0 or itogo != 0):
                itogo_mismatch.append([r, name, b["month"], bn, nal, itogo])
            if bn != 0 or nal != 0:
                d = by_level.setdefault(level, {}).setdefault(b["month"], [0.0, 0.0, 0])
                d[0] += bn
                d[1] += nal
                d[2] += 1
        if name is None and not any_val:
            continue
        if name is not None:
            last_named_row = r
        row = {"r": r, "level": level, "bold": bold, "indent": indent, "hidden": hidden_row,
               "name": name, "name_col": name_col, "vals": vals}
        info["rows"].append(row)
        if level == 2 and name is not None:
            names_lvl2.append(name)

    info["by_level"] = {str(k): v for k, v in by_level.items()}
    info["negatives"] = negatives
    info["itogo_mismatch"] = itogo_mismatch
    info["last_named_row"] = last_named_row

    print(f"  строк с именем/значением: {len(info['rows'])}, последняя именованная строка: {last_named_row}")
    for level in sorted(by_level):
        for m in sorted(by_level[level]):
            bn, nal, cnt = by_level[level][m]
            print(f"    уровень {level}  {m}: безнал {fmt(bn):>18}  нал {fmt(nal):>16}  ненулевых ячеек-строк {cnt}")
    lv01 = [(x["r"], x["level"], x["bold"], x["name"]) for x in info["rows"] if x["level"] in (0, 1) and x["name"]]
    print(f"  уровни 0/1 (r, lvl, bold, имя): {lv01[:40]}")
    print(f"  уровень 2, имён: {len(names_lvl2)}; первые: {names_lvl2[:12]}")
    if negatives:
        print(f"  ОТРИЦАТЕЛЬНЫЕ ({len(negatives)}): {negatives[:8]}")
    if itogo_mismatch:
        print(f"  Итого <> Безнал+Нал ({len(itogo_mismatch)}): {itogo_mismatch[:6]}")
    comments = [(x["r"], x["name"], m, v["comment"]) for x in info["rows"] for m, v in x["vals"].items() if v["comment"]]
    if comments:
        print(f"  комментарии ({len(comments)}): {comments[:10]}")
    # жирность vs уровень
    bold_by_level = {}
    for x in info["rows"]:
        if x["name"]:
            bold_by_level.setdefault(x["level"], [0, 0])[1 if x["bold"] else 0] += 1
    print(f"  жирность по уровням {{уровень: [не жирных, жирных]}}: {bold_by_level}")
    dump["sheets"].append(info)

print("\n" + "#" * 100)
print(f"ИТОГО листов {len(wb.sheetnames)}, скрытых {hidden_cnt}, видимых {len(wb.sheetnames) - hidden_cnt}")
print("Порядок листов:", [(s.title, s.sheet_state) for s in wb.worksheets])

with open(OUT_JSON, "w", encoding="utf-8") as f:
    json.dump(dump, f, ensure_ascii=False, indent=1, default=str)
print("Дамп:", OUT_JSON)
