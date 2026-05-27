"""
Preflight: изучить эталон А_ОтражениеЗПпоКазне №000000006.

EXPECTED (зафиксировано 2026-05-27):
- ОтражениеЗПпоКазне UID: eb4bdb3e-59d1-11f1-a2ec-ec86c6e33afc
- ТЧ.РаспределениеКазна: 408 строк (из них 210 с пустым ДокРаспределениеЗП — пропуск)
- Уникальных Ф2 (из непустых): 7
  - №000000025: 106 строк, Σ=716,000.00
  - №000000028: 58 строк, Σ=1,638,200.00
  - №000000023: 1 строк, Σ=23,800.00
  - №000000032: 6 строк, Σ=69,700.00
  - №000000026: 21 строк, Σ=348,800.00 (уже есть ВКассу 000Ц-000002)
  - №000000027: 2 строк, Σ=12,127.00
  - №000000029: 4 строк, Σ=145,500.00
- Σ всех Ф2-строк: 2,954,127.00
- Строк с А_РаспределениеЗаработнойПлаты: 0
- А_Привилегированный.ОбеспечитьПодразделенияОрганизацийПоПодразделению доступен из COM ✓

Smoke-тест должен создать 7 ВКассу (для №000000026 — обновить существующую 000Ц-000002).
"""
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)

# Найти эталон
q = erp.NewObject("Запрос")
q.Текст = '''ВЫБРАТЬ ПЕРВЫЕ 1 О.Ссылка КАК Ссылка
ИЗ Документ.А_ОтражениеЗПпоКазне КАК О
ГДЕ О.Номер = "000000006" И О.Проведен'''
rs = q.Выполнить()
if rs.Пустой():
    print("FAIL: А_ОтражениеЗПпоКазне №000000006 не найдено")
    sys.exit(1)
sel = rs.Выбрать(); sel.Следующий()
otr_ref = sel.Ссылка
print(f"ОтражениеЗПпоКазне: UID={erp.string(otr_ref.УникальныйИдентификатор())}")

otr = otr_ref.ПолучитьОбъект()
print(f"  Дата: {otr.Дата} | Организация: {otr.Организация.Наименование}")
print(f"  ТЧ.РаспределениеКазна строк: {otr.РаспределениеКазна.Количество()}")

# Группировка по ДокРаспределениеЗП
unique_f2 = {}
type_skip = 0
empty_doc = 0
total_sum = 0.0

for i in range(otr.РаспределениеКазна.Количество()):
    r = otr.РаспределениеКазна.Получить(i)
    doc = r.ДокРаспределениеЗП
    if not doc or doc.Пустая():
        empty_doc += 1
        continue
    meta_name = str(doc.Метаданные().Имя)
    if meta_name != "РаспределениеФ2":
        type_skip += 1
        continue
    f2_uid = erp.string(doc.УникальныйИдентификатор())
    if f2_uid not in unique_f2:
        unique_f2[f2_uid] = {"ref": doc, "rows": 0, "sum": 0.0, "num": str(doc.Номер)}
    unique_f2[f2_uid]["rows"] += 1
    unique_f2[f2_uid]["sum"] += float(r.Сумма or 0)
    total_sum += float(r.Сумма or 0)

print(f"\n  Уникальных Ф2: {len(unique_f2)}")
print(f"  Строк с пустым ДокРаспределениеЗП: {empty_doc}")
print(f"  Строк типа А_РаспределениеЗаработнойПлаты (пропускаем): {type_skip}")
print(f"  Σ Сумма (только Ф2-строки): {total_sum:,.2f}")
for uid, info in unique_f2.items():
    print(f"    Ф2 №{info['num']}: {info['rows']} строк, Σ={info['sum']:,.2f}, UID={uid}")

# Уникальные пары (Сотр, Подр) с СтатьейДДС
print("\n  Уникальные (Сотрудник / Подразделение) → СтатьяДДС:")
pairs_seen = set()
pairs_with_empty_statya = 0
for i in range(otr.РаспределениеКазна.Количество()):
    r = otr.РаспределениеКазна.Получить(i)
    if not r.ДокРаспределениеЗП or r.ДокРаспределениеЗП.Пустая():
        continue
    if str(r.ДокРаспределениеЗП.Метаданные().Имя) != "РаспределениеФ2":
        continue
    sotr = str(r.Сотрудник) if r.Сотрудник and not r.Сотрудник.Пустая() else "—"
    podr = str(r.Подразделение) if r.Подразделение and not r.Подразделение.Пустая() else "—"
    statya = str(r.СтатьяДвиженияДенежныхСредств) if r.СтатьяДвиженияДенежныхСредств and not r.СтатьяДвиженияДенежныхСредств.Пустая() else "—(пусто)—"
    pair = f"{sotr}|{podr}"
    if pair in pairs_seen:
        continue
    pairs_seen.add(pair)
    if statya == "—(пусто)—":
        pairs_with_empty_statya += 1
    print(f"    {sotr} / {podr} → {statya}")

print(f"\n  Уникальных пар (Сотр, Подр): {len(pairs_seen)}")
print(f"  Пар с пустой СтатьейДДС: {pairs_with_empty_statya}")

# Проверка доступности А_Привилегированный из Python COM
print("\n=== Проверка доступа к А_Привилегированный из Python COM ===")
try:
    test_подр = None
    for i in range(otr.РаспределениеКазна.Количество()):
        r = otr.РаспределениеКазна.Получить(i)
        if r.Подразделение and not r.Подразделение.Пустая():
            test_подр = r.Подразделение
            break
    if test_подр is not None:
        # Доступ через атрибут с кириллическим именем
        try:
            mod = getattr(erp, "А_Привилегированный")
            podr_org = mod.ОбеспечитьПодразделенияОрганизацийПоПодразделению(test_подр)
            print(f"  ✓ А_Привилегированный доступен. Конвертация {test_подр} → {podr_org}")
        except Exception as e_inner:
            info = e_inner.excepinfo[2] if hasattr(e_inner, "excepinfo") and e_inner.excepinfo else str(e_inner)
            print(f"  ✗ getattr() упало: {info}")
    else:
        print("  WARN: нет ни одной строки с заполненным Подразделением для проверки")
except Exception as e:
    info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
    print(f"  ✗ внешняя ошибка: {info}")

print("\nDONE — фиксируй значения для использования в последующих тестах")
