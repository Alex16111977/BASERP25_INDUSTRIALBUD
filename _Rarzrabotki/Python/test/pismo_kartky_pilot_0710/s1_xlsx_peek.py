import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")
B = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Письмо Директору по отчетам на производство\07102026"
for p in ("IRC 2026", "МД ООН Админ"):
    wb = openpyxl.load_workbook(B + "\\" + p + "\План-факт виконання СС на 07-10-2026.xlsx", data_only=True)
    ws = wb.active
    print("==", p, ws.title, ws.max_row, ws.max_column)
    for r in range(1, 40):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        ind = ws.cell(r, 1).alignment.indent
        if any(v is not None for v in vals):
            print(r, ind, vals)
