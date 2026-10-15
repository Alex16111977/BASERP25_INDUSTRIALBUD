# -*- coding: utf-8 -*-
"""Распознавание товарных строк счетов: xls/xlsx/docx/doc/rtf/pdf(текст) → RAB/<uid>/stroki.json.

Строка товара — строка таблицы, где есть название и тройка чисел «количество, цена, сумма» с
количество × цена ≈ сумма (или × 1,2 — цена без ПДВ, сумма с ПДВ). Из найденных строк выбирается блок, итог которого
совпадает с суммой счёта Казны (конкурентный лист: несколько поставщиков — берётся блок нужной суммы).
Сканы (jpg/png, pdf без текста) получают sposob="изображение" и строки не заполняются — их распознаёт ИИ.

    python raspoznat.py            # все счета
    python raspoznat.py <uid> ...  # выборочно
"""
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')
RAB = os.environ.get("TOVARY_RAB", r"C:\Razrabotki\tovary_schetov_kazna")
UNITS = {"шт", "шт.", "м", "м.", "м2", "м²", "м3", "м.п", "м.п.", "мп", "п.м", "пм", "кг", "т", "л", "уп", "уп.",
         "компл", "компл.", "к-т", "комп", "пар", "рул", "рул.", "лист", "лист.", "пач", "пач.", "ящ", "мішок", "меш",
         "послуга", "посл", "посл.", "грн", "год", "м.кв", "кв.м", "м/п", "бух", "пог.м", "пог. м", "тн", "набір"}
NUM = re.compile(r"^-?\d[\d\s\u00a0]*([.,]\d+)?$")


def num(s):
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    t = str(s).strip().replace("\u00a0", " ")
    if not t or not NUM.match(t):
        return None
    t = t.replace(" ", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


# ---------------- чтение в «сетку» строк ----------------
def grid_xls(path):
    import xlrd
    b = xlrd.open_workbook(path)
    out = []
    for sh in b.sheets():
        for r in range(sh.nrows):
            out.append(list(sh.row_values(r)))
    return out


def resave_excel(path):
    """xlsx без sharedStrings (выгрузка 1С) → пересохранить через Excel COM."""
    dst = os.path.splitext(path)[0] + ".conv.xlsx"
    if os.path.exists(dst):
        return dst
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path), ReadOnly=True)
        wb.SaveAs(os.path.abspath(dst), FileFormat=51)
        wb.Close(False)
    finally:
        xl.Quit()
    return dst


def grid_xlsx(path):
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except KeyError:
        wb = openpyxl.load_workbook(resave_excel(path), data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            out.append(list(row))
    return out


def grid_docx(path):
    import docx
    d = docx.Document(path)
    out = []
    for t in d.tables:
        for row in t.rows:
            cells, prev = [], None
            for c in row.cells:
                if c._tc is prev:      # объединённые ячейки повторяются
                    continue
                prev = c._tc
                cells.append(c.text.strip())
            out.append(cells)
    return out


def to_docx(path):
    """doc/rtf → docx через Word COM (рядом, тем же именем)."""
    dst = os.path.splitext(path)[0] + ".conv.docx"
    if os.path.exists(dst):
        return dst
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    try:
        doc = w.Documents.Open(os.path.abspath(path), ReadOnly=True)
        doc.SaveAs2(os.path.abspath(dst), FileFormat=16)
        doc.Close(False)
    finally:
        w.Quit()
    return dst


def grid_pdf(path):
    """Таблицы PyMuPDF; если их нет — строки текста, разбитые по 2+ пробелам. Возвращает (grid, текст_есть)."""
    import pymupdf
    doc = pymupdf.open(path)
    out, chars = [], 0
    for page in doc:
        txt = page.get_text()
        chars += len(txt.strip())
        try:
            tabs = page.find_tables()
            for t in tabs.tables:
                for row in t.extract():
                    out.append([("" if c is None else str(c).replace("\n", " ")) for c in row])
        except Exception:
            pass
        for line in page.get_text("text").splitlines():
            parts = [p for p in re.split(r"\s{2,}|\t", line.strip()) if p]
            if len(parts) >= 3:
                out.append(parts)
    return out, chars >= 50


# ---------------- строки товаров ----------------
def parse_row(cells):
    vals = [("" if c is None else (c if isinstance(c, (int, float)) else str(c).strip())) for c in cells]
    nums = [(i, num(v)) for i, v in enumerate(vals)]
    nums = [(i, v) for i, v in nums if v is not None]
    texts = [(i, str(v)) for i, v in enumerate(vals) if num(v) is None and str(v).strip()]
    if len(nums) < 3 or not texts:
        return None
    best = None
    for a in range(len(nums)):
        for b in range(a + 1, len(nums)):
            for c in range(b + 1, len(nums)):
                (ia, q), (ib, p), (ic, s) = nums[a], nums[b], nums[c]
                if q <= 0 or p <= 0 or s <= 0:
                    continue
                for k in (1.0, 1.2):
                    if abs(q * p * k - s) <= max(0.06, 0.005 * s):
                        cand = (ic - ia, ia, q, p, s, k)
                        if best is None or cand[5] < best[5] or (cand[5] == best[5] and ic > best[1]):
                            best = cand
    if not best:
        return None
    _, ia, q, p, s, k = best
    name_cells = [t for i, t in texts if t.lower().strip(" .") not in UNITS and len(t) > 2 and i < ia]
    if not name_cells:
        name_cells = [t for i, t in texts if t.lower().strip(" .") not in UNITS and len(t) > 2]
    if not name_cells:
        return None
    name = max(name_cells, key=len)
    unit = next((t for i, t in texts if t.lower().strip() in UNITS), "")
    if re.fullmatch(r"(разом|всього|итого|усього|сума|сумма).*", name.lower()):
        return None
    return dict(nazv=re.sub(r"\s+", " ", name).strip(), kol=round(q, 3), ed=unit.strip(),
                cena=round(p * k, 2), summa=round(s, 2))


def rows_from_grid(grid):
    out, seen = [], set()
    for cells in grid:
        r = parse_row(cells)
        if r:
            key = (r["nazv"], r["kol"], r["summa"])
            if key in seen:           # pdf: строки из таблицы и из текста дублируются
                continue
            seen.add(key)
            out.append(r)
    return out


def pick_block(rows, total):
    """Все строки, если итог совпал; иначе непрерывный блок с итогом = сумме счёта (конкурентный лист).
    Строки без ПДВ при сумме счёта с ПДВ (итог × 1,2 = сумма счёта) переводятся в суммы с ПДВ."""
    if not rows:
        return [], "строки не найдены"
    for k in (1.0, 1.2):
        if abs(round(sum(r["summa"] for r in rows), 2) * k - total) <= 1:
            return scale(rows, k), ""
    for k in (1.0, 1.2):
        for i in range(len(rows)):
            acc = 0.0
            for j in range(i, len(rows)):
                acc += rows[j]["summa"]
                if abs(acc * k - total) <= 1:
                    return scale(rows[i:j + 1], k), ""
    s = round(sum(r["summa"] for r in rows), 2)
    return rows, f"сума рядків {s:.2f} ≠ сума рахунку {total:.2f}"


def scale(rows, k):
    if k == 1.0:
        return rows
    out = []
    for r in rows:
        x = dict(r)
        x["cena"] = round(r["cena"] * k, 2)
        x["summa"] = round(r["summa"] * k, 2)
        x["bez_pdv"] = True
        out.append(x)
    return out


def recognize(schet):
    best = None
    for f in schet["files"]:
        if not f.get("path"):
            continue
        ext, path = f["ext"], f["path"]
        try:
            if ext == "xls":
                grid, sposob = grid_xls(path), "excel"
            elif ext == "xlsx":
                grid, sposob = grid_xlsx(path), "excel"
            elif ext == "docx":
                grid, sposob = grid_docx(path), "word"
            elif ext in ("doc", "rtf"):
                grid, sposob = grid_docx(to_docx(path)), "word"
            elif ext == "pdf":
                grid, has_text = grid_pdf(path)
                sposob = "pdf-текст" if has_text else "изображение"
                if not has_text:
                    grid = []
            else:
                grid, sposob = [], "изображение"
        except Exception as e:
            grid, sposob = [], f"ошибка: {e}"[:60]
        rows = rows_from_grid(grid)
        block, warn = pick_block(rows, schet["summa_scheta"])
        res = dict(uid=schet["uid"], file=os.path.basename(path), sposob=sposob, stroki=block, predupr=warn)
        score = (0 if not warn and block else 1 if block else 2)
        if best is None or score < best[0]:
            best = (score, res)
        if score == 0:
            break
    if best is None:
        return dict(uid=schet["uid"], file="", sposob="нет файла", stroki=[], predupr="файла нет")
    return best[1]


def main():
    scheta = json.load(open(os.path.join(RAB, "scheta.json"), encoding="utf-8"))
    only = set(sys.argv[1:])
    stat = {}
    for s in scheta:
        if only and s["uid"] not in only:
            continue
        out = os.path.join(RAB, s["uid"], "stroki.json")
        if os.path.exists(out) and json.load(open(out, encoding="utf-8")).get("sposob") == "ии":
            stat["ии (не трогаем)"] = stat.get("ии (не трогаем)", 0) + 1
            continue
        res = recognize(s)
        for i, r in enumerate(res["stroki"], 1):
            r["n"] = i
        os.makedirs(os.path.dirname(out), exist_ok=True)
        json.dump(res, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        key = ("OK " if res["stroki"] and not res["predupr"] else "РАСХОЖД " if res["stroki"] else "НЕТ ") + res["sposob"]
        stat[key] = stat.get(key, 0) + 1
    for k, v in sorted(stat.items()):
        print(f"{v:4d}  {k}")


if __name__ == "__main__":
    main()
