# -*- coding: utf-8 -*-
# Заповнюємо КодКазна для ФізОсіб де А_КазнаСоздан=True але КодКазна порожній
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
conn_kazn = v8.Connect('Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"')
S = conn.String
Sk = conn_kazn.String

# Казна: ФІО → Код
qk = conn_kazn.NewObject("Запрос")
qk.Text = "ВЫБРАТЬ С.Наименование КАК ФИО, С.Код КАК Код, С.ИНН КАК ИНН ИЗ Справочник.Сотрудники КАК С ГДЕ НЕ С.ПометкаУдаления"
sk = qk.Execute().Choose()
kazna = {}
while sk.Next():
    fio = Sk(sk.ФИО).strip()
    if fio:
        kazna[fio] = {"kod": Sk(sk.Код).strip(), "inn": Sk(sk.ИНН).strip()}

# ERP: ФізОсоби з А_КазнаСоздан=True без КодКазна
q = conn.NewObject("Запрос")
q.Text = "ВЫБРАТЬ Ф.Ссылка, Ф.Наименование КАК ФИО, Ф.А_КодСотрудникаКазна КАК КодК ИЗ Справочник.ФизическиеЛица КАК Ф ГДЕ Ф.А_КазнаСоздан = ИСТИНА"
r = q.Execute().Choose()
while r.Next():
    fio = S(r.ФИО).strip()
    kodK = S(r.КодК).strip()
    if not kodK and fio in kazna:
        obj = r.Ссылка.GetObject()
        obj.А_КодСотрудникаКазна = kazna[fio]["kod"]
        obj.ОбменДанными.Загрузка = True
        obj.Записать()
        print(f"Updated: {fio} → КодКазна={kazna[fio]['kod']}")
    else:
        print(f"OK: {fio}, КодКазна={kodK}")
