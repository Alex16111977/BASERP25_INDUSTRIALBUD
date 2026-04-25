"""Создать один тестовый документ и проверить, сохраняется ли ТЧ."""
import sys, datetime
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp, uuid_str

c = connect_erp()
podr = c.Справочники.СтруктураПредприятия.НайтиПоНаименованию("Астарта. Тищенки", True)
org = c.Справочники.Организации.НайтиПоНаименованию('ТОВ "ІНДАСТРІАЛБУД"', True)
# any existing article
q = c.NewObject("Запрос")
q.Текст = "ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.А_Статьи_PL ГДЕ НЕ ЭтоГруппа И НЕ ПометкаУдаления"
tz = q.Выполнить().Выгрузить()
art = tz.Получить(0).Ссылка
print("Test article:", art.Наименование)

def test(with_load_flag):
    doc = c.Документы.А_ОтчетPL.СоздатьДокумент()
    doc.Дата = datetime.datetime(2025, 12, 15, 12, 0, 0)
    doc.Организация = org
    doc.Подразделение = podr
    doc.ПодразделениеСтрока = "TEST " + ("с Загрузка=Истина" if with_load_flag else "без Загрузка")
    row = doc.ДанныеОтчета.Добавить()
    row.Статья = art
    row.СуммаФома1 = 100.50
    row.СуммаФорма2 = 200.75
    row.Сумма = 301.25
    row.Комментарий = "тест"

    if with_load_flag:
        doc.ОбменДанными.Загрузка = True
    try:
        doc.Записать()
        u = uuid_str(c, doc.Ссылка)
        # Now read back ТЧ
        q2 = c.NewObject("Запрос")
        q2.Текст = ("ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК N, СУММА(Сумма) КАК С "
                    "ИЗ Документ.А_ОтчетPL.ДанныеОтчета ГДЕ Ссылка = &Р")
        q2.УстановитьПараметр("Р", doc.Ссылка)
        res = q2.Выполнить().Выгрузить().Получить(0)
        return f"OK uuid={u} rows={int(res.N)} sumaТЧ={float(res.С):.2f}"
    except Exception as e:
        return f"ERR: {e}"

print("\n--- Тест 1: БЕЗ ОбменДанными.Загрузка ---")
print(" ", test(False))
print("\n--- Тест 2: С ОбменДанными.Загрузка=Истина ---")
print(" ", test(True))
