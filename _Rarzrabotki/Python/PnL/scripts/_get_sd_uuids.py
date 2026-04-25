"""Get UUIDs of needed СтатьиДоходов."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp, uuid_str

c = connect_erp()
q = c.NewObject("Запрос")
q.Текст = ("ВЫБРАТЬ СД.Ссылка, СД.Наименование ИЗ ПланВидовХарактеристик.СтатьиДоходов КАК СД "
           "ГДЕ СД.Наименование В (\"Прочие доходи\", \"Прочие доходы от Фин деятельности\")")
tz = q.Выполнить().Выгрузить()
for i in range(tz.Количество()):
    r = tz.Получить(i)
    print(f"{r.Наименование}  →  {uuid_str(c, r.Ссылка)}")
