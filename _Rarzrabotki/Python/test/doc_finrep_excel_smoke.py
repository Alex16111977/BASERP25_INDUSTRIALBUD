# -*- coding: utf-8 -*-
"""Smoke частини 2: документ на 25.09 перераховується в пам'яті (не записується), ДвоичныеДанныеExcel() →
xlsx; Excel через COM відкриває файл, перераховує формули; K кожного рядка договору = документ,
аркуш «25.09.2026», закріплення до рядка 5, примітки є, у договорах ОРІЛЬ/IRC — строки оплати."""
import os, sys, tempfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
buh = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
S = buh.String
провалів = 0


def перевірка(умова, текст):
    global провалів
    print(("  ✓ " if умова else "  ✗ ") + текст)
    провалів += 0 if умова else 1


q = buh.NewObject("Запрос")
q.Текст = ("ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК С ИЗ Документ.А_ФинансовыйОтчетПроизводства КАК Д "
           "ГДЕ НЕ Д.ПометкаУдаления И Д.ДатаРасчета МЕЖДУ ДАТАВРЕМЯ(2026,9,25) И ДАТАВРЕМЯ(2026,9,25,23,59,59)")
док = q.Execute().Выгрузить().Получить(0).С.ПолучитьОбъект()
р = док.Рассчитать()
перевірка(bool(р.Выполнено) and int(р.Ошибок) == 0, f"Рассчитать() на 25.09: рядків {док.Расчет.Количество()}")
коди = {S(док.Контроль.Получить(i).КодЗамечания) for i in range(док.Контроль.Количество())}
перевірка("К24" in коди, "К24 — сховані договори ООН 2025 показано в контролі")
K_док = [float(док.Расчет.Получить(i).ИтогоКПолучению) for i in range(док.Расчет.Количество())
         if S(док.Расчет.Получить(i).Договор) != ""]

шлях = os.path.join(tempfile.gettempdir(), "finrep_2509_smoke.xlsx")
док.ДвоичныеДанныеExcel().Записать(шлях)
перевірка(os.path.getsize(шлях) > 5000, f"xlsx записано ({os.path.getsize(шлях)} байт)")

xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    кн = xl.Workbooks.Open(шлях)
    xl.CalculateFull()
    ар = кн.Worksheets(1)
    перевірка(ар.Name == "25.09.2026", f"аркуш «{ар.Name}»")
    ар.Activate()
    перевірка(xl.ActiveWindow.FreezePanes and xl.ActiveWindow.SplitRow == 5, "закріплено до рядка 5")
    перевірка(ар.Range("A4").Value == "Контрагент" and ар.Range("K4").Value == "ИТОГО сумма к получению", "шапка")
    K_xl, формули_ок, примітки_строк = [], True, 0
    останній = ар.UsedRange.Row + ар.UsedRange.Rows.Count
    for r in range(7, останній + 1):
        if ар.Rows(r).OutlineLevel == 2:
            K_xl.append(float(ар.Cells(r, 11).Value or 0))
            формули_ок &= str(ар.Cells(r, 7).Formula).startswith("=E")
    for к in ар.Comments:
        т = к.Text()
        if "оплата до" in т or "прострочено" in т or "строк оплати" in т:
            примітки_строк += 1
    перевірка(len(K_xl) == len(K_док), f"рядків договорів: Excel {len(K_xl)}, документ {len(K_док)}")
    розбіжності = [(i, a, b) for i, (a, b) in enumerate(zip(K_xl, K_док)) if abs(a - b) > 0.02]
    перевірка(not розбіжності, f"K після перерахунку Excel = документ (розбіжностей {len(розбіжності)}: {розбіжності[:3]})")
    перевірка(формули_ок, "G договорів — формула =E−F")
    перевірка(ар.Comments.Count > 20, f"приміток {ар.Comments.Count}")
    перевірка(примітки_строк >= 2, f"примітки зі строком оплати: {примітки_строк}")
    кн.Close(False)
finally:
    xl.Quit()
print(f"\nпровалів: {провалів}")
