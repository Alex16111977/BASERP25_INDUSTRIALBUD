# -*- coding: utf-8 -*-
"""Верификация канона А_Статьи_PL.Сорт: порядок статей каждой группы в справочнике
(УПОРЯДОЧИТЬ ПО Сорт) должен начинаться со статей Excel-листа «Глобино-2» (июнь-2026)
в порядке листа; не-Excel статьи — строго после них."""
import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')
import openpyxl

XLSX = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Червень_26\!PL по компании Червень 2026.xlsx"
# Диапазоны статей листа по группам (эталон Глобино-2)
RANGES = {
    "000000006": range(6, 8),      # ОД
    "000000001": range(9, 34),     # СС
    "000000007": range(36, 37),    # ДР
    "000000003": range(42, 49),    # ОПЗ
    "000000005": range(51, 53),    # МЗ
    "000000002": range(55, 72),    # АЗ
    "000000008": range(74, 75),    # НС
    "000000004": [78, 79, 82, 85], # ФД: ФинДоход, ФинРасходы, Налог, Дивиденды
}

wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
ws = wb["Глобино-2"]
cells = {}
for idx, row in enumerate(ws.iter_rows(min_col=2, max_col=2, max_row=95), start=1):
    v = row[0].value
    cells[idx] = v.strip() if isinstance(v, str) else ""
excel = {code: [cells[r] for r in rng if cells.get(r)] for code, rng in RANGES.items()}

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ Ст.Наименование КАК Наименование, Ст.Группа.Код КАК КодГруппы, Ст.Сорт КАК Сорт
ИЗ Справочник.А_Статьи_PL КАК Ст
ГДЕ НЕ Ст.ПометкаУдаления
УПОРЯДОЧИТЬ ПО Ст.Группа.Сорт, Ст.Сорт, Ст.Код"""
t = q.Execute().Выгрузить()
cat = {}
for i in range(t.Количество()):
    r = t.Получить(i)
    cat.setdefault(S(r.КодГруппы), []).append((S(r.Наименование).strip(), float(r.Сорт)))

fails = 0
for code, excel_list in excel.items():
    db_list = cat.get(code, [])
    db_names = [n.lower() for n, _ in db_list]
    # префикс базы = Excel-список (по порядку)
    prefix = db_names[:len(excel_list)]
    exp = [e.lower() for e in excel_list]
    if prefix == exp:
        extra = [n for n, _ in db_list[len(excel_list):]]
        print(f"OK  группа {code}: {len(excel_list)} Excel-статей в порядке листа"
              + (f", хвост: {extra}" if extra else ""))
    else:
        fails += 1
        print(f"FAIL группа {code}:")
        for i, e in enumerate(exp):
            got = prefix[i] if i < len(prefix) else "<нет>"
            mark = "  " if got == e else "!!"
            print(f"  {mark} лист: {e!r:60s} база: {got!r}")
print("ИТОГ:", "OK" if fails == 0 else f"FAIL x{fails}")
sys.exit(1 if fails else 0)
