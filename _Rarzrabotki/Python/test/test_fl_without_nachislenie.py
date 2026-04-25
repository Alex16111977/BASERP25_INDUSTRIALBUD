# -*- coding: utf-8 -*-
# Знайти ФізОсіб які є в НДФЛ/ВзносыФОТ/Удержания, АЛЕ немає в Начисления.
# Саме такі ФізОсоби, ймовірно, блокують проведення.
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

# Перезаписати НЗ
q0 = erp.NewObject("Запрос")
q0.Text = 'ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка ИЗ Документ.А_ОтражениеЗПпоКазне КАК Д ГДЕ Д.Номер = "000000001"'
sel0 = q0.Execute().Выбрать()
sel0.Следующий()
sel0.Ссылка.ПолучитьОбъект().СоздатьДокументНачисления()

q = erp.NewObject("Запрос")
q.Text = (
    'ВЫБРАТЬ ПЕРВЫЕ 1 НЗ.Ссылка '
    'ИЗ Документ.НачислениеЗарплаты КАК НЗ '
    'ГДЕ НЗ.А_ДокОтражениеЗПпоКазне.Номер = "000000001" И НЕ НЗ.ПометкаУдаления '
    'УПОРЯДОЧИТЬ ПО НЗ.Дата УБЫВ'
)
sel = q.Execute().Выбрать()
sel.Следующий()
нз = sel.Ссылка.ПолучитьОбъект()
print(f"НЗ: {S(sel.Ссылка)}")

# Зібрати множини ФізОсіб
nach_fl = {}
for i in range(нз.Начисления.Количество()):
    r = нз.Начисления.Получить(i)
    try:
        uid = S(r.Сотрудник.ФизическоеЛицо.УникальныйИдентификатор())
        nach_fl[uid] = S(r.Сотрудник.ФизическоеЛицо)
    except:
        pass

ud_fl = {}
for i in range(нз.Удержания.Количество()):
    r = нз.Удержания.Получить(i)
    try:
        uid = S(r.ФизическоеЛицо.УникальныйИдентификатор())
        ud_fl[uid] = S(r.ФизическоеЛицо)
    except:
        pass

ndfl_fl = {}
for i in range(нз.НДФЛ.Количество()):
    r = нз.НДФЛ.Получить(i)
    try:
        uid = S(r.ФизическоеЛицо.УникальныйИдентификатор())
        ndfl_fl[uid] = S(r.ФизическоеЛицо)
    except:
        pass

vzn_fl = {}
for i in range(нз.ВзносыФОТ.Количество()):
    r = нз.ВзносыФОТ.Получить(i)
    try:
        uid = S(r.ФизическоеЛицо.УникальныйИдентификатор())
        vzn_fl[uid] = S(r.ФизическоеЛицо)
    except:
        pass

print(f"\nУнікальних ФізОсіб:")
print(f"  Начисления:     {len(nach_fl)}")
print(f"  Удержания:      {len(ud_fl)}")
print(f"  НДФЛ:           {len(ndfl_fl)}")
print(f"  ВзносыФОТ:      {len(vzn_fl)}")

# НДФЛ без Начисления
ndfl_no_nach = {uid: v for uid, v in ndfl_fl.items() if uid not in nach_fl}
print(f"\n=== У НДФЛ АЛЕ НЕМАЄ в Начисления: {len(ndfl_no_nach)} ===")
for uid, v in list(ndfl_no_nach.items())[:15]:
    print(f"  {v}")

# ВзносыФОТ без Начисления
vzn_no_nach = {uid: v for uid, v in vzn_fl.items() if uid not in nach_fl}
print(f"\n=== У ВзносыФОТ АЛЕ НЕМАЄ в Начисления: {len(vzn_no_nach)} ===")
for uid, v in list(vzn_no_nach.items())[:15]:
    print(f"  {v}")

# Удержания без Начисления
ud_no_nach = {uid: v for uid, v in ud_fl.items() if uid not in nach_fl}
print(f"\n=== У Удержания АЛЕ НЕМАЄ в Начисления: {len(ud_no_nach)} ===")
for uid, v in list(ud_no_nach.items())[:15]:
    print(f"  {v}")
