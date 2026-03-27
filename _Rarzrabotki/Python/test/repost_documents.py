# -*- coding: utf-8 -*-
"""
Перепроведение документов после удаления РегистраторовРасчетов.
"""
import win32com.client
import sys

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("OK")

# Получаем enum через свойство подключения
mode_unpost = conn.РежимЗаписиДокумента.ОтменаПроведения
mode_post = conn.РежимЗаписиДокумента.Проведение
print(f"Write modes OK")

doc_types = [
    "ПоступлениеБезналичныхДенежныхСредств",
    "СписаниеБезналичныхДенежныхСредств",
    "ПриходныйКассовыйОрдер",
    "РасходныйКассовыйОрдер",
    "АвансовыйОтчет",
    "ПриобретениеТоваровУслуг",
    "РеализацияТоваровУслуг",
    "ВозвратТоваровОтКлиента",
    "ВозвратТоваровПоставщику",
    "ВводОстатков",
]

query = conn.NewObject("Запрос")
total_reposted = 0
total_errors = 0
all_errors = []

for dt in doc_types:
    print(f"\n--- {dt} ---")
    query.Text = (
        "ВЫБРАТЬ Ссылка, Номер "
        "ИЗ Документ." + dt + " "
        "ГДЕ Проведен = ИСТИНА "
        "И Дата >= ДАТАВРЕМЯ(2025,12,1) "
        "И Дата < ДАТАВРЕМЯ(2026,4,1) "
        "УПОРЯДОЧИТЬ ПО Дата"
    )
    try:
        result = query.Execute().Выбрать()
    except Exception as e:
        print(f"  Skip: {str(e)[:80]}")
        continue

    count = 0
    errors = 0
    while result.Следующий():
        try:
            doc_obj = result.Ссылка.ПолучитьОбъект()
            doc_obj.Записать(mode_unpost)
            doc_obj.Записать(mode_post)
            count += 1
            if count % 100 == 0:
                print(f"  Reposted: {count}...")
        except Exception as e:
            errors += 1
            err_str = str(e)[:120]
            if errors <= 3:
                print(f"  ERROR {result.Номер}: {err_str}")
            all_errors.append(f"{dt} {result.Номер}: {err_str}")

    print(f"  Result: {count} OK, {errors} errors")
    total_reposted += count
    total_errors += errors

print(f"\n{'='*60}")
print(f"TOTAL: reposted {total_reposted}, errors {total_errors}")

# Check
query.Text = "ВЫБРАТЬ КОЛИЧЕСТВО(Ссылка) КАК Кол ИЗ Документ.РегистраторРасчетов"
result = query.Execute().Выбрать()
result.Следующий()
print(f"New registrators: {result.Кол}")

# Check 00DL-7185
query.Text = """
    ВЫБРАТЬ РАЗЛИЧНЫЕ Регистратор КАК Рег
    ИЗ РегистрНакопления.РасчетыСПоставщикамиПоСрокам
    ГДЕ ДокументРегистратор.Номер = "00DL-7185"
      И ДокументРегистратор.Дата = ДАТАВРЕМЯ(2025,12,10)
"""
try:
    result = query.Execute().Выбрать()
    cnt = 0
    while result.Следующий():
        cnt += 1
    print(f"Registrators for 00DL-7185: {cnt} (was 3)")
except:
    pass

if all_errors:
    print(f"\nFirst 20 errors:")
    for e in all_errors[:20]:
        print(f"  {e}")
