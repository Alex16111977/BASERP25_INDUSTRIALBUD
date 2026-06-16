# -*- coding: utf-8 -*-
"""Rule #-1 пред-загрузочный тест ПРАВКИ 1 (CASE ФормаPL в запросе реквизитов РКО).
Извлекает изменённый запрос ЗаполнитьПараметрыИнициализации ИЗ ФАЙЛА, прогоняет для N0000053062,
проверяет: исполняется + колонка ФормаPL = Форма2 (ОргБух пуст)."""
import io, sys
import win32com.client
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

BSL = r"C:\Configuration_downloads\BASERP25\Documents\РасходныйКассовыйОрдер\Ext\ManagerModule.bsl"
content = io.open(BSL, encoding="utf-8-sig").read()  # text mode → \n

def depipe(lit):
    out = []
    for i, ln in enumerate(lit.split("\n")):
        out.append(ln if i == 0 else (ln[ln.find("|")+1:] if "|" in ln else ln))
    return "\n".join(out).replace('""', '"')

# извлечь запрос реквизитов: в процедуре ЗаполнитьПараметрыИнициализации, после 'Запрос.Текст ='
i = content.index("Процедура ЗаполнитьПараметрыИнициализации")
j = content.index("Запрос.Текст =", i)
op = content.index('"', j)          # открывающая кавычка запроса
cl = content.index('";', op + 1)    # закрывающая
rekviz = depipe(content[op+1:cl])

erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# найти документ
qd = erp.NewObject("Запрос")
qd.Text = "ВЫБРАТЬ Ссылка ИЗ Документ.РасходныйКассовыйОрдер ГДЕ Номер = &Н"
qd.SetParameter("Н", "N0000053062")
sel = qd.Execute().Выбрать(); sel.Следующий(); doc = sel.Ссылка

q = erp.NewObject("Запрос")
q.Text = rekviz
q.SetParameter("Ссылка", doc)
try:
    r = q.Execute().Выгрузить()
except Exception as e:
    info = getattr(e, 'excepinfo', None)
    print("❌ Запрос реквизитов НЕ исполнился:", info[2] if info else e); sys.exit(1)

# есть ли колонка ФормаPL?
cols = [c.Имя for c in r.Колонки]
print("Колонка 'ФормаPL' в запросе реквизитов:", "ЕСТЬ" if "ФормаPL" in cols else "НЕТ ❌")
if "ФормаPL" in cols and r.Количество() > 0:
    val = r.Получить(0).ФормаPL
    pred = erp.XMLСтрока(val)
    print(f"✅ Исполнилось. ФормаPL для N0000053062 = {pred}  (ожидаем Форма2, т.к. А_ОрганизацияБухгалтерия пуст)")
    print("ВЕРДИКТ:", "✅ OK" if pred == "Форма2" else f"⚠️ неожидаемо: {pred}")
