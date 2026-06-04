# -*- coding: utf-8 -*-
"""Быстрый дамп реквизитов + движений Списание/ПостБезнал."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

def check(doc_type, number):
    print("\n" + "=" * 100)
    print(f"{doc_type}  Номер={number}")
    print("=" * 100)
    q = erp.NewObject("Запрос")
    q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.{doc_type} ГДЕ Номер = "{number}"'
    sel = q.Выполнить().Выбрать()
    if not sel.Следующий():
        print("[FAIL] не найден"); return
    ref = sel.Ссылка
    print(f"Документ: {S(ref)}")
    obj = ref.ПолучитьОбъект()
    for fld in ("Дата","Номер","Проведен","Организация","Подразделение","Касса","БанковскийСчет",
                "ХозяйственнаяОперация","Валюта","ВалютаКонвертации","СуммаДокумента",
                "А_ОбработанКазна","А_ВведенВЕРП","Партнер","Контрагент"):
        try:
            v = getattr(obj, fld)
            if isinstance(v,(str,int,float,bool)): print(f"  {fld:<24}: {v}")
            else: print(f"  {fld:<24}: {S(v) if v and erp.ЗначениеЗаполнено(v) else '(пусто)'}")
        except: pass

    # Движения в РНДС.ВПути
    print("\n--- РНДС.ВПути ---")
    q.УстановитьПараметр("Док", ref)
    q.Текст = "ВЫБРАТЬ Вп.* ИЗ РегистрНакопления.ДенежныеСредстваВПути КАК Вп ГДЕ Вп.Регистратор = &Док"
    r = q.Выполнить().Выгрузить()
    print(f"Строк: {r.Количество()}")
    for i in range(r.Количество()):
        rec = r.Получить(i)
        vd = S(getattr(rec,"ВидДвижения",""))
        podr_v = getattr(rec,"Подразделение",None)
        podr = S(podr_v) if podr_v and erp.ЗначениеЗаполнено(podr_v) else "(пусто)"
        s_ = getattr(rec,"Сумма",0)
        get_v = getattr(rec,"Получатель",None); otp_v = getattr(rec,"Отправитель",None)
        get_s = S(get_v) if get_v and erp.ЗначениеЗаполнено(get_v) else "—"
        otp_s = S(otp_v) if otp_v and erp.ЗначениеЗаполнено(otp_v) else "—"
        vp_v = getattr(rec,"ВидПереводаДенежныхСредств",None)
        vp = S(vp_v) if vp_v and erp.ЗначениеЗаполнено(vp_v) else ""
        print(f"  {vd:<8} Σ={s_:>14}  Подр={podr:<22}  {vp[:25]:<27}  {otp_s} → {get_s}")

    # Движения в РСКПС/РСППС (для понимания риска COM-репоста)
    for reg in ("РасчетыСКлиентамиПоСрокам", "РасчетыСПоставщикамиПоСрокам"):
        q.Текст = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К ИЗ РегистрНакопления.{reg} ГДЕ Регистратор = &Док"
        rr = q.Выполнить().Выгрузить()
        if rr.Количество() > 0 and int(rr.Получить(0).К) > 0:
            print(f"  ⚠ Есть {int(rr.Получить(0).К)} движений в {reg} — COM-репост ОПАСЕН (потеря РегистраторРасчётов)")

check("СписаниеБезналичныхДенежныхСредств", "00000019546")
check("ПоступлениеБезналичныхДенежныхСредств", "00DL-007179")
