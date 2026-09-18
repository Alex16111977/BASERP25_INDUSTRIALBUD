# -*- coding: utf-8 -*-
"""§3.3 промта: проверить, что Excel COM Rows(r).OutlineLevel даёт те же уровни, что openpyxl.

Книга открывается ReadOnly, ничего не пишется. Сравнение по 5 листам, строки 1..200.
Excel нумерует уровни с 1 (1 = без группировки), openpyxl — с 0. Ожидание: COM = openpyxl + 1.
Заодно: Visible скрытых листов, тип .Value датовых ячеек строки 3, IndentLevel/Bold колонки A.
"""
import sys, datetime
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client
from openpyxl import load_workbook

PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Бюджет прогноз\Отчеты высланные\ДДС Производство Вересень - Грудень 2026 Прогноз.xlsx"
SHEETS = ["МД IRC 2026", "ЦО_Производство", "МД ООН 2026", "СВОД_Производство", "МК Підгірці"]

xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(PATH, False, True)   # UpdateLinks=False, ReadOnly=True
try:
    print("Excel:", xl.Version, "листов:", wb.Sheets.Count)
    wbo = load_workbook(PATH, data_only=True)
    total = 0
    bad = 0
    for name in SHEETS:
        ws = wb.Sheets(name)
        wso = wbo[name]
        mism = []
        for r in range(1, 201):
            lv_com = ws.Rows(r).OutlineLevel
            lv_op = wso.row_dimensions[r].outline_level or 0
            total += 1
            if lv_com != lv_op + 1:
                mism.append((r, lv_com, lv_op))
        bad += len(mism)
        print(f"«{name}» Visible={ws.Visible}: расхождений COM≠openpyxl+1 в строках 1..200: {len(mism)} {mism[:10]}")
        print("   COM  6..30:", [ws.Rows(r).OutlineLevel for r in range(6, 31)])
        print("   opxl 6..30:", [wso.row_dimensions[r].outline_level or 0 for r in range(6, 31)])
    print(f"ИТОГО строк сравнено {total}, расхождений {bad}")
    # видимость всех листов
    vis = [(wb.Sheets(i).Name, wb.Sheets(i).Visible) for i in range(1, wb.Sheets.Count + 1)]
    print("Visible по листам (-1 видим, 0 скрыт, 2 очень скрыт):", vis)
    ws = wb.Sheets("МД IRC 2026")
    print("IRC строка 3, .Value по колонкам 1,2,6,10,14,15,16:")
    for c in (1, 2, 6, 10, 14, 15, 16):
        v = ws.Cells(3, c).Value
        print(f"   c{c}: {type(v).__name__} {v!r}")
    print("IRC колонка A: (ряд, значение, IndentLevel, Bold, OutlineLevel)")
    for r in (6, 7, 8, 15, 19, 20, 21, 26, 27, 37, 123, 155, 156):
        a = ws.Cells(r, 1)
        print(f"   r{r}: {a.Value!r} indent={a.IndentLevel} bold={a.Font.Bold} lvl={ws.Rows(r).OutlineLevel} hidden={ws.Rows(r).Hidden}")
    print("IRC r19 значения c2..c5:", [ws.Cells(19, c).Value for c in range(2, 6)])
    print("IRC r8 значения c2..c5:", [ws.Cells(8, c).Value for c in range(2, 6)])
finally:
    wb.Close(False)
    xl.Quit()
print("готово")
