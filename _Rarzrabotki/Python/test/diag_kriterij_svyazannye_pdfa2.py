# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def err(e):
    info = getattr(e, 'excepinfo', None)
    return info[2] if info and len(info) > 2 and info[2] else str(e)

def crit(ref):
    q = erp.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ СД.Ссылка КАК Ссылка ИЗ КритерийОтбора.СвязанныеДокументы(&З) КАК СД"
    q.SetParameter("З", ref)
    r = q.Execute().Выгрузить()
    out = []
    for i in range(r.Количество()):
        out.append(r.Получить(i).Ссылка)
    return out

def s(x):
    try:
        return erp.String(x)
    except Exception:
        return "<ref>"

# целевой документ из скрина
q0 = erp.NewObject("Запрос")
q0.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК Реф ИЗ Документ.А_ПриходДенегОтФинАгента ГДЕ Номер = &Н"
q0.SetParameter("Н", "N00023780")
r0 = q0.Execute().Выгрузить()
if r0.Количество() > 0:
    ref = r0.Получить(0).Реф
    print("=== N00023780 ===")
    try:
        ch = crit(ref)
        print("дети:", len(ch))
        for c in ch:
            print("  ->", s(c))
            try:
                ch2 = crit(c)
                print("     внуки:", len(ch2))
            except Exception as e:
                print("     ВНУК-КРИТЕРИЙ FAIL:", s(c), "::", err(e))
    except Exception as e:
        print("N00023780 КРИТЕРИЙ FAIL:", err(e))
else:
    print("N00023780 не найден")

# скан партии: ищем документы, на которых критерий падает
print("=== скан 80 документов ===")
qs = erp.NewObject("Запрос")
qs.Text = "ВЫБРАТЬ ПЕРВЫЕ 80 Ссылка КАК Реф, Номер КАК Ном ИЗ Документ.А_ПриходДенегОтФинАгента ГДЕ НЕ ПометкаУдаления УПОРЯДОЧИТЬ ПО Дата"
rs = qs.Execute().Выгрузить()
fails = 0
withkids = 0
for i in range(rs.Количество()):
    row = rs.Получить(i)
    try:
        ch = crit(row.Реф)
        if len(ch) > 0:
            withkids += 1
            # рекурсия 1 уровень
            for c in ch:
                try:
                    crit(c)
                except Exception as e:
                    fails += 1
                    print("FAIL on child of", row.Ном, "child=", s(c), "::", err(e))
    except Exception as e:
        fails += 1
        print("FAIL on", row.Ном, "::", err(e))
print("итог: с детьми=%d, падений=%d" % (withkids, fails))
print("DONE")
