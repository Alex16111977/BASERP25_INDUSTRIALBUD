# -*- coding: utf-8 -*-
"""Найти РасчетКурсовыхРазниц + дамп движений в РНДС.ВПути."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

# Все РасчетКурсовыхРазниц которые писали в РНДС.ВПути
q = erp.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ РАЗЛИЧНЫЕ Вп.Регистратор КАК Док,
    ПРЕДСТАВЛЕНИЕ(Вп.Регистратор) КАК Имя,
    КОЛИЧЕСТВО(*) КАК Движений,
    СУММА(Вп.Сумма) КАК Σ
ИЗ РегистрНакопления.ДенежныеСредстваВПути КАК Вп
ГДЕ ТИПЗНАЧЕНИЯ(Вп.Регистратор) = ТИП(Документ.РасчетКурсовыхРазниц)
СГРУППИРОВАТЬ ПО Вп.Регистратор
УПОРЯДОЧИТЬ ПО Вп.Регистратор.Дата
"""
r = q.Выполнить().Выгрузить()
print(f"РасчетКурсовыхРазниц с движениями в РНДС.ВПути: {r.Количество()}\n")
for i in range(r.Количество()):
    rec = r.Получить(i)
    print(f"  {rec.Имя}  Движ={rec.Движений}  Σ={rec.Σ}")

if r.Количество() == 0:
    print("Нет документов"); sys.exit(0)

# Берём первый — детально
ref = r.Получить(0).Док
print(f"\n=== Детально: {S(ref)} ===")
obj = ref.ПолучитьОбъект()
md = obj.Метаданные()
print("\n--- Реквизиты шапки ---")
for r in md.Реквизиты:
    name = r.Имя
    try:
        v = getattr(obj, name)
        if v is None or v == "" or v == 0 or v is False: continue
        if isinstance(v,(str,int,float,bool)): print(f"  {name}: {v}")
        else:
            try: print(f"  {name}: {S(v) if erp.ЗначениеЗаполнено(v) else '(пусто)'}")
            except: pass
    except: pass

# Движения в РНДС.ВПути (детально)
print(f"\n--- Движения в РНДС.ВПути ---")
q2 = erp.NewObject("Запрос")
q2.УстановитьПараметр("Док", ref)
q2.Текст = "ВЫБРАТЬ Вп.* ИЗ РегистрНакопления.ДенежныеСредстваВПути КАК Вп ГДЕ Вп.Регистратор = &Док"
rr = q2.Выполнить().Выгрузить()
print(f"Строк: {rr.Количество()}")
cols = [c.Имя for c in rr.Колонки]
for i in range(rr.Количество()):
    rec = rr.Получить(i)
    vd = S(getattr(rec,"ВидДвижения",""))
    podr_v = getattr(rec,"Подразделение",None)
    podr = S(podr_v) if podr_v and erp.ЗначениеЗаполнено(podr_v) else "(пусто)"
    s_ = getattr(rec,"Сумма",0)
    get_v = getattr(rec,"Получатель",None); otp_v = getattr(rec,"Отправитель",None)
    get_s = S(get_v) if get_v and erp.ЗначениеЗаполнено(get_v) else "—"
    otp_s = S(otp_v) if otp_v and erp.ЗначениеЗаполнено(otp_v) else "—"
    vp_v = getattr(rec,"ВидПереводаДенежныхСредств",None)
    vp = S(vp_v) if vp_v and erp.ЗначениеЗаполнено(vp_v) else ""
    val_v = getattr(rec,"Валюта",None)
    val = S(val_v) if val_v and erp.ЗначениеЗаполнено(val_v) else ""
    print(f"  {vd:<8} Σ={s_:>10}  {val:<5} Подр={podr:<22}  {vp[:25]:<27}  {otp_s} → {get_s}")
