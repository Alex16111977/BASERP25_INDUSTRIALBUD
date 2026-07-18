"""
Discovery: структура РегистрНакопления.БДДС у BuhKazn (Казна).

Цель: задокументировать реальные имена реквизитов БДДС, которые использует
ЗагрузитьРаспределениеКазна_ОтражениеЗП() — Подразделение, Направление, Сотрудник,
Регистратор.Организация.ЕГРПОУ, СтатьяДвиженияДенежныхСредств, Сумма, СуммаНачисления,
СуммаФОТ, СуммаНалогов.

Запуск: C:\\Python313\\python.exe discovery_kazna_bdds_struktura.py
"""
import io
import sys
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

ARTIFACT_DIR = os.path.join(os.path.dirname(__file__), "..", "_artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)

OUT_FILE = os.path.join(ARTIFACT_DIR, "kazna_bdds_schema.md")


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    kazna = v8.Connect('Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"')

    # Метаданные регистра — список реквизитов/измерений/ресурсов
    try:
        md = kazna.Метаданные.РегистрыНакопления.Найти("БДДС")
    except Exception as e:
        print(f"FAIL metadata: {e}")
        return

    lines = []
    lines.append("# РегНакопл.БДДС (BuhKazn) — структура\n")
    lines.append(f"Discovery: {sys.argv[0]}\n")
    lines.append("\n## Измерения\n")
    for d in md.Измерения:
        lines.append(f"- `{d.Имя}` — {d.Тип}\n")
    lines.append("\n## Ресурсы\n")
    for r in md.Ресурсы:
        lines.append(f"- `{r.Имя}` — {r.Тип}\n")
    lines.append("\n## Реквизиты\n")
    for a in md.Реквизиты:
        lines.append(f"- `{a.Имя}` — {a.Тип}\n")

    # Test query — взять 5 строк из любого доступного периода 2026
    q = kazna.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 5
    БДДС.Подразделение.Код КАК ПодразделениеКод,
    БДДС.Направление.Код КАК НаправлениеКод,
    БДДС.Сотрудник.ИНН КАК СотрудникИНН,
    БДДС.Сотрудник.Наименование КАК СотрудникФИО,
    БДДС.Регистратор КАК Регистратор,
    БДДС.Регистратор.Дата КАК ДатаДок,
    БДДС.СтатьяДвиженияДенежныхСредств.Код КАК СтатьяКод,
    БДДС.Сумма КАК Сумма,
    БДДС.СуммаНачисления КАК СуммаНачисления,
    БДДС.СуммаФОТ КАК СуммаФОТ,
    БДДС.СуммаНалогов КАК СуммаНалогов
ИЗ
    РегистрНакопления.БДДС КАК БДДС
ГДЕ
    (БДДС.Регистратор ССЫЛКА Документ.РаспределениеЗаработнойПлаты
     ИЛИ БДДС.Регистратор ССЫЛКА Документ.РаспределениеФ2)
УПОРЯДОЧИТЬ ПО БДДС.Период УБЫВ
"""
    try:
        rs = q.Выполнить()
        ts = rs.Выгрузить()
        rows_count = ts.Количество()
        lines.append(f"\n## Тестовый запрос — найдено {rows_count} строк\n\n")
        if rows_count > 0:
            lines.append("| ПодрКод | НаправлКод | ИНН | ФИО | СтатьяКод | Сумма | СумНачисл | СумФОТ | СумНалогов |\n")
            lines.append("|---|---|---|---|---|---:|---:|---:|---:|\n")
            for i in range(min(5, rows_count)):
                r = ts.Получить(i)
                lines.append(
                    f"| {r.ПодразделениеКод} | {r.НаправлениеКод} | {r.СотрудникИНН} | "
                    f"{r.СотрудникФИО} | {r.СтатьяКод} | {r.Сумма} | {r.СуммаНачисления} | "
                    f"{r.СуммаФОТ} | {r.СуммаНалогов} |\n"
                )
            # Извлечь UUID одного регистратора — для verify_uuid_lookup_kazna_erp.py
            uuid_str = kazna.string(ts.Получить(0).Регистратор.УникальныйИдентификатор())
            lines.append(f"\n**Образец UUID регистратора[0]:** `{uuid_str}` (тип: {type(ts.Получить(0).Регистратор).__name__})\n")
    except Exception as e:
        info = ""
        if hasattr(e, "excepinfo") and e.excepinfo:
            info = e.excepinfo[2]
        lines.append(f"\n**ERROR test query:** {info or e}\n")

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.writelines(lines)
    print(f"OK -> {OUT_FILE}")


if __name__ == "__main__":
    main()
