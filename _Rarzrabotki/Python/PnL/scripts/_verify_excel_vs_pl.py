"""Проверка #1: Excel-итоги ↔ Документ.А_ОтчетPL (ИтогиПоГруппам).

Для каждого документа А_ОтчетPL проверяет:
- Сумма_Excel (из парсинга Excel) = Сумма_Расчёт (сумма детальных строк ДанныеОтчета)
- Расхождение = 0

Результат → data/reports/verification_excel_vs_pl.csv

Формат CSV:
    Дата;Подразделение;Группа;СуммаExcel;СуммаРасчёт;Расхождение;Статус
"""
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp

OUT_CSV = config.REPORTS_DIR / "verification_excel_vs_pl.csv"


QUERY = """
ВЫБРАТЬ
    Док.Дата КАК Дата,
    Док.Номер КАК Номер,
    Док.Подразделение.Наименование КАК Подразделение,
    ИГ.Группа.Наименование КАК Группа,
    ИГ.Сумма_Excel КАК СуммаExcel,
    ИГ.Сумма_Расчёт КАК СуммаРасчёт,
    ИГ.Расхождение КАК Расхождение
ИЗ Документ.А_ОтчетPL КАК Док
ЛЕВОЕ СОЕДИНЕНИЕ Документ.А_ОтчетPL.ИтогиПоГруппам КАК ИГ
    ПО ИГ.Ссылка = Док.Ссылка
ГДЕ НЕ Док.ПометкаУдаления
УПОРЯДОЧИТЬ ПО Док.Дата, Док.Подразделение.Наименование, ИГ.НомерСтроки
"""


def main():
    conn = connect_erp()
    q = conn.NewObject("Запрос")
    q.Текст = QUERY
    tz = q.Выполнить().Выгрузить()
    n = tz.Количество()
    print(f"Строк в ИтогиПоГруппам: {n}")

    config.REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    with OUT_CSV.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Дата", "Номер", "Подразделение", "Группа",
                    "СуммаExcel", "СуммаРасчёт", "Расхождение", "Статус"])
        total = 0
        total_ok = 0
        total_err = 0
        for i in range(n):
            r = tz.Получить(i)
            дата = str(r.Дата)[:10] if r.Дата else ""
            номер = str(r.Номер) if r.Номер else ""
            подр = str(r.Подразделение) if r.Подразделение else ""
            групп = str(r.Группа) if r.Группа else ""
            sx = float(r.СуммаExcel or 0)
            sr = float(r.СуммаРасчёт or 0)
            расхожд = float(r.Расхождение or 0)
            статус = "OK" if abs(расхожд) < 0.01 else "РАСХОЖДЕНИЕ"
            total += 1
            if статус == "OK":
                total_ok += 1
            else:
                total_err += 1
            w.writerow([дата, номер, подр, групп,
                        f"{sx:.2f}", f"{sr:.2f}", f"{расхожд:.2f}", статус])

    print(f"Всего строк: {total}")
    print(f"  OK (Расхождение=0): {total_ok}")
    print(f"  РАСХОЖДЕНИЕ: {total_err}")
    print(f"Результат: {OUT_CSV}")


if __name__ == "__main__":
    main()
