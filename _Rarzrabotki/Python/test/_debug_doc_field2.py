# -*- coding: utf-8 -*-
"""Проверка составного поля Документ (план=А_СтруктураСебестоимости, факт=ПриобретениеТоваровУслуг)."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def run_batch(text):
    q = erp.NewObject("Запрос")
    q.МенеджерВременныхТаблиц = erp.NewObject("МенеджерВременныхТаблиц")
    q.Текст = text
    return q.Выполнить().Выгрузить()

FULL = open(r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\_doc_field_query.txt", encoding="utf-8").read()
t = run_batch(FULL)
names = [t.Колонки.Получить(j).Имя for j in range(t.Колонки.Количество())]
idx = {names[j]: j for j in range(len(names))}
n = t.Количество()
sСС = sФакт = 0.0
типПлан = типФакт = None
for i in range(n):
    r = t.Получить(i)
    сс = float(r.Получить(idx["СуммаСС"]) or 0)
    сум = float(r.Получить(idx["Сумма"]) or 0)
    sСС += сс; sФакт += сум
    док = r.Получить(idx["Документ"])
    тип = док.Метаданные().Имя if (док is not None and hasattr(док, "Метаданные")) else None
    if сс and not типПлан and тип: типПлан = тип
    if сум and not типФакт and тип: типФакт = тип
print(f"строк={n}  Σплан={sСС:.2f}  Σфакт={sФакт:.2f}")
print(f"тип Документа у ПЛАНовой строки:   {типПлан}")
print(f"тип Документа у ФАКТической строки: {типФакт}")
ok = abs(sСС - 5699881.55) < 1 and sФакт > 5900000 and типПлан == "А_СтруктураСебестоимости" and типФакт == "ПриобретениеТоваровУслуг"
print("OK: поле Документ корректно, суммы не изменились." if ok else "ВНИМАНИЕ: проверь значения.")
