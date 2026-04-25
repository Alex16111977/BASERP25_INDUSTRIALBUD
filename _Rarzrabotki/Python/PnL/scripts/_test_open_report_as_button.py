"""E2E-тест: симулирует кнопку 'Открыть сверку с 1С' и выгружает итоговые настройки отчёта в XML,
чтобы увидеть реально применённые Период/Отбор/ПараметрыДанных.
"""
import sys
import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp

OUT_XML = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\_rasshifrovka_settings.xml"


def xs(conn, val):
    try:
        return str(conn.XMLСтрока(val))
    except Exception as e:
        return f"<error: {e}>"


conn = connect_erp()
rep = conn.Отчеты.А_ОтчетPL.Создать()
схема = rep.СхемаКомпоновкиДанных

вариант = схема.ВариантыНастроек.Найти("А_ОтчетPL_РасшифровкаДокументаОтчетPL")
if вариант is None:
    print("FAIL: вариант не найден"); sys.exit(1)
print(f"Вариант: {вариант.Имя}")

rep.КомпоновщикНастроек.ЗагрузитьНастройки(вариант.Настройки)

Период = conn.NewObject("СтандартныйПериод")
Период.ДатаНачала = datetime.datetime(2025, 12, 1)
Период.ДатаОкончания = datetime.datetime(2025, 12, 31, 23, 59, 59)
Подразд = conn.Справочники.СтруктураПредприятия.НайтиПоНаименованию("Крушинка", True)

for i in range(rep.КомпоновщикНастроек.ПользовательскиеНастройки.Элементы.Количество()):
    э = rep.КомпоновщикНастроек.ПользовательскиеНастройки.Элементы.Получить(i)
    id_ = str(э.ИдентификаторПользовательскойНастройки)
    if id_ == "Период":
        э.Значение = Период
        э.Использование = True
    elif id_ == "Подразделение":
        э.ПравоеЗначение = Подразд
        э.Использование = True

итог = rep.КомпоновщикНастроек.ПолучитьНастройки()

print("\n=== Отбор (детально) ===")
for i in range(итог.Отбор.Элементы.Количество()):
    э = итог.Отбор.Элементы.Получить(i)
    left_xml = xs(conn, э.ЛевоеЗначение)
    vid = str(э.ВидСравнения)
    right_xml = xs(conn, э.ПравоеЗначение)
    print(f"[{i}] Левое={left_xml!r}")
    print(f"    ВидСравнения={vid}")
    print(f"    Правое={right_xml[:200]!r}")
    print(f"    Использование={э.Использование}")

print("\n=== ПараметрыДанных (детально) ===")
for i in range(итог.ПараметрыДанных.Элементы.Количество()):
    э = итог.ПараметрыДанных.Элементы.Получить(i)
    try:
        param_name = xs(conn, э.Параметр)
    except Exception:
        param_name = "?"
    try:
        value_xml = xs(conn, э.Значение)
    except Exception:
        value_xml = "<err>"
    print(f"[{i}] {param_name}  Значение={value_xml[:200]!r}  Исп={э.Использование}")

# Выгрузим итоговые настройки в XML файл
try:
    xml_zapis = conn.NewObject("ЗаписьXML")
    xml_zapis.УстановитьСтроку()
    conn.СериализаторXDTO.ЗаписатьXML(xml_zapis, итог)
    xml_text = str(xml_zapis.Закрыть())
    with open(OUT_XML, "w", encoding="utf-8") as f:
        f.write(xml_text)
    print(f"\nНастройки XML: {OUT_XML}  (size={len(xml_text)})")
except Exception as e:
    print(f"Ошибка выгрузки XML: {e}")

# Попробуем выполнить запрос отчёта напрямую, имитируя ПолучитьОбъединенныеДанные (копия SQL)
print("\n=== Симуляция запроса отчёта за декабрь 2025 ===")
import textwrap
q_text = textwrap.dedent("""
ВЫБРАТЬ
	ТЧ.Ссылка.Подразделение КАК Подразделение,
	СУММА(ТЧ.Сумма) КАК СуммаPL
ИЗ Документ.А_ОтчетPL.ДанныеОтчета КАК ТЧ
ГДЕ ТЧ.Ссылка.Дата МЕЖДУ &НП И &КП И НЕ ТЧ.Ссылка.ПометкаУдаления
СГРУППИРОВАТЬ ПО ТЧ.Ссылка.Подразделение
""").strip()
q = conn.NewObject("Запрос")
q.Текст = q_text
q.УстановитьПараметр("НП", datetime.datetime(2025, 12, 1))
q.УстановитьПараметр("КП", datetime.datetime(2025, 12, 31, 23, 59, 59))
tz = q.Выполнить().Выгрузить()
n = tz.Количество()
print(f"Подразделений в документах А_ОтчетPL за декабрь 2025: {n}")
for i in range(n):
    r = tz.Получить(i)
    подр = str(r.Подразделение) if r.Подразделение else ""
    сумма = float(r.СуммаPL) if r.СуммаPL else 0.0
    marker = "  ← Крушинка" if подр == "Крушинка" else ""
    print(f"  {подр[:40]:40} | СуммаPL={сумма:>14.2f}{marker}")
