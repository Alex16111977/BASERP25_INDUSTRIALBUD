# -*- coding: utf-8 -*-
"""Быстрый дамп реквизитов + движений РКО N0000053020."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

q = erp.NewObject("Запрос")
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.РасходныйКассовыйОрдер ГДЕ Номер = "N0000053020"'
sel = q.Выполнить().Выбрать()
if not sel.Следующий():
    print("[FAIL] РКО N0000053020 не найден"); sys.exit(1)
ref = sel.Ссылка
print(f"Документ: {S(ref)}")
obj = ref.ПолучитьОбъект()
for fld in ("Дата","Номер","Проведен","Организация","Подразделение","Касса","ХозяйственнаяОперация",
            "Валюта","ВалютаКонвертации","СуммаДокумента","А_ОбработанКазна","А_ВведенВЕРП"):
    try:
        v = getattr(obj, fld)
        if isinstance(v,(str,int,float,bool)): print(f"  {fld:<22}: {v}")
        else: print(f"  {fld:<22}: {S(v) if v and erp.ЗначениеЗаполнено(v) else '(пусто)'}")
    except: pass

# Касса.Подразделение
try:
    pk = obj.Касса.Подразделение
    print(f"\n  Касса.Подразделение: {S(pk) if erp.ЗначениеЗаполнено(pk) else '(пусто)'}")
except: pass

# Движения в РНДС.ВПути
print("\n=== Движения в РНДС.ВПути ===")
q.УстановитьПараметр("Док", ref)
q.Текст = """
ВЫБРАТЬ Вп.* ИЗ РегистрНакопления.ДенежныеСредстваВПути КАК Вп ГДЕ Вп.Регистратор = &Док
"""
r = q.Выполнить().Выгрузить()
print(f"Строк: {r.Количество()}")
cols = [c.Имя for c in r.Колонки]
for i in range(r.Количество()):
    rec = r.Получить(i)
    vd = S(getattr(rec,"ВидДвижения",""))
    podr_v = getattr(rec,"Подразделение",None)
    podr = S(podr_v) if podr_v and erp.ЗначениеЗаполнено(podr_v) else "(пусто)"
    s_ = getattr(rec,"Сумма",0)
    get_v = getattr(rec,"Получатель",None)
    otp_v = getattr(rec,"Отправитель",None)
    get_s = S(get_v) if get_v and erp.ЗначениеЗаполнено(get_v) else "—"
    otp_s = S(otp_v) if otp_v and erp.ЗначениеЗаполнено(otp_v) else "—"
    print(f"  {vd:<8} Σ={s_:>10}  Подр={podr:<22}  {otp_s} → {get_s}")
