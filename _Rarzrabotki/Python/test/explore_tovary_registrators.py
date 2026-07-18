# -*- coding: utf-8 -*-
"""Розвідка: які типи документів рухають РН.ТоварыНаСкладах (гістограма типів)."""
import win32com.client, sys, datetime

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

q = erp.NewObject("Запрос")
q.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 400 "
    "  Дв.Регистратор КАК Рег "
    "ИЗ РегистрНакопления.ТоварыНаСкладах КАК Дв "
    "УПОРЯДОЧИТЬ ПО Дв.Период УБЫВ")
tz = q.Execute().Unload()

hist = {}        # тип -> [кількість, приклад_ref, проведен]
for i in range(tz.Count()):
    reg = tz.Get(i).Рег
    try:
        nm = reg.Метаданные().Имя
    except Exception:
        nm = "?"
    if nm not in hist:
        hist[nm] = [0, reg, None]
        try:
            hist[nm][2] = reg.Проведен
        except Exception:
            pass
    hist[nm][0] += 1

print(f"Усього семпл рядків: {tz.Count()}; типів: {len(hist)}")
for nm, (cnt, ref, posted) in sorted(hist.items(), key=lambda kv: -kv[1][0]):
    print(f"  {cnt:4d}  {nm:45s} проведен_прикладу={posted}  приклад={S(ref)}")
print("=== DONE ===")
