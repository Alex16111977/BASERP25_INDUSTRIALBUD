# -*- coding: utf-8 -*-
# Диагностика: воспроизводим запрос КритерийОтбора.СвязанныеДокументы для А_ПриходДенегОтФинАгента
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def err(e):
    info = getattr(e, 'excepinfo', None)
    return info[2] if info and len(info) > 2 and info[2] else str(e)

# 1) ref документа
q0 = erp.NewObject("Запрос")
q0.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК Реф, Номер КАК Ном ИЗ Документ.А_ПриходДенегОтФинАгента ГДЕ НЕ ПометкаУдаления УПОРЯДОЧИТЬ ПО Дата УБЫВ"
r0 = q0.Execute().Выгрузить()
if r0.Количество() == 0:
    print("НЕТ ДОКУМЕНТОВ"); sys.exit()
ref = r0.Получить(0).Реф
print("doc:", r0.Получить(0).Ном)

# 2) КритерийОтбора.СвязанныеДокументы(наш док) — кто ссылается на наш документ
q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ СД.Ссылка КАК Ссылка ИЗ КритерийОтбора.СвязанныеДокументы(&З) КАК СД"
q.SetParameter("З", ref)
children = []
try:
    r = q.Execute().Выгрузить()
    print("КРИТЕРИЙ(нашДок) OK, детей =", r.Количество())
    for i in range(r.Количество()):
        c = r.Получить(i).Ссылка
        children.append(c)
        try:
            print("  ->", erp.String(c))
        except Exception:
            print("  -> <ref>")
except Exception as e:
    print("КРИТЕРИЙ(нашДок) FAIL:", err(e))

# 3) Рекурсия: критерий для КАЖДОГО дочернего (как делает форма СвязанныеДокументы рекурсивно)
print("--- рекурсия по детям ---")
for c in children[:30]:
    q2 = erp.NewObject("Запрос")
    q2.Text = "ВЫБРАТЬ СД.Ссылка КАК Ссылка ИЗ КритерийОтбора.СвязанныеДокументы(&З) КАК СД"
    q2.SetParameter("З", c)
    try:
        r2 = q2.Execute().Выгрузить()
        cm = ""
        try:
            cm = erp.String(c)
        except Exception:
            cm = "<ref>"
        print("  child КРИТЕРИЙ OK (%d):" % r2.Количество(), cm)
    except Exception as e:
        cm = "<ref>"
        try:
            cm = erp.String(c)
        except Exception:
            pass
        print("  child КРИТЕРИЙ FAIL:", cm, "::", err(e))

print("DONE")
