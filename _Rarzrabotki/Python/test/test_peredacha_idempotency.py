# -*- coding: utf-8 -*-
"""
IDEMPOTENCY: два прогона СоздатьПередачуНачисленийМеждуПодразделениями() на одном
родителе → ТОТ ЖЕ документ (Find-Or-Create) и идентичные ТЧ/движения (Δ=0).
Запуск: C:\\Python313\\python.exe <этот файл>
"""
import win32com.client as w
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

erp = w.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')


def run(text, **p):
    z = erp.NewObject("Запрос"); z.Текст = text
    for k, v in p.items():
        z.УстановитьПараметр(k, v)
    return z.Выполнить().Выгрузить()


def scalar(text, **p):
    tz = run(text, **p)
    return getattr(tz.Получить(0), "Р") if tz.Количество() else None


parent = run('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Документ.А_ОтражениеЗПпоКазне ГДЕ Номер = "000000005" И Проведен').Получить(0).С


def прогон_и_снимок():
    parent.ПолучитьОбъект().СоздатьПередачуНачисленийМеждуПодразделениями()
    nd = run('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Документ.А_ПередачаНачисленийМеждуПодразделениями ГДЕ А_ДокументОснование = &Р', Р=parent).Получить(0).С
    uuid = erp.string(nd.УникальныйИдентификатор())
    s = {}
    for тч in ("Начисления", "Налоги", "Удержания"):
        row = run(f'ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К, ЕСТЬNULL(СУММА(Сумма),0) КАК С ИЗ Документ.А_ПередачаНачисленийМеждуПодразделениями.{тч} ГДЕ Ссылка = &Д', Д=nd).Получить(0)
        s[тч] = (row.К, round(float(row.С), 2))
    pap = run('ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К, ЕСТЬNULL(СУММА(П.Сумма),0) КАК С ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК П ГДЕ П.Регистратор = &Д', Д=nd).Получить(0)
    s["ПАП"] = (pap.К, round(float(pap.С), 2))
    vzs = run('ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К, ЕСТЬNULL(СУММА(ВзСС.СуммаВзаиморасчетов),0) КАК С ИЗ РегистрНакопления.А_ВзаиморасчетыССотрудниками КАК ВзСС ГДЕ ВзСС.Регистратор = &Д', Д=nd).Получить(0)
    s["А_ВзСС"] = (vzs.К, round(float(vzs.С), 2))
    return uuid, s


uuid1, s1 = прогон_и_снимок()
uuid2, s2 = прогон_и_снимок()

print("UUID прогон1:", uuid1)
print("UUID прогон2:", uuid2)
for k in s1:
    print(f"  {k:<12} прогон1={s1[k]}  прогон2={s2[k]}  {'OK' if s1[k]==s2[k] else 'DIFF!'}")

fails = []
if uuid1 != uuid2:
    fails.append(f"разные документы: {uuid1} != {uuid2} (Find-Or-Create не сработал)")
for k in s1:
    if s1[k] != s2[k]:
        fails.append(f"{k}: {s1[k]} != {s2[k]} (не идемпотентно)")

print("\n" + ("IDEMPOTENCY PASS" if not fails else "IDEMPOTENCY FAIL:\n  - " + "\n  - ".join(fails)))
sys.exit(0 if not fails else 1)
