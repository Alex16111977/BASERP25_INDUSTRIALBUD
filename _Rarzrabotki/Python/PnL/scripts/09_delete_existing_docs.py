"""Pometka удаления всех документов А_ОтчетPL для указанных периодов (помечает на удаление)."""
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp


def main():
    conn = connect_erp()
    for f in config.EXCEL_FILES:
        period_date = datetime.strptime(f["period"], "%Y-%m-%d")
        month_start = period_date.replace(day=1, hour=0, minute=0, second=0)
        month_end = period_date.replace(hour=23, minute=59, second=59)
        q = conn.NewObject("Запрос")
        q.Текст = (
            "ВЫБРАТЬ Ссылка ИЗ Документ.А_ОтчетPL "
            "ГДЕ Дата МЕЖДУ &С И &ПО И НЕ ПометкаУдаления"
        )
        q.УстановитьПараметр("С", month_start)
        q.УстановитьПараметр("ПО", month_end)
        tz = q.Выполнить().Выгрузить()
        n = tz.Количество()
        print(f"{f['label']}: {n} документ(ов) за {f['period']}")
        for i in range(n):
            ref = tz.Получить(i).Ссылка
            obj = ref.ПолучитьОбъект()
            obj.УстановитьПометкуУдаления(True)
        if n:
            print(f"  помечены на удаление")


if __name__ == "__main__":
    main()
