"""
Discovery: структура регистров zup_1 и zup_2 — реальных источников налогов
и начислений для документа А_ОтражениеЗПпоКазне.

Обращаем внимание: в PROMPT.md источник назван «BuhBud (BAS Бухгалтерія)», но
реальный код ObjectModule.bsl подключается к **zup_1** + **zup_2**:
  - zup_1: РегНакопл.ВзаиморасчетыПоНДФЛ + РегРасчет.ВзносыВФонды +
           РегРасчет.УдержанияРаботниковОрганизаций + Документ.ИсполнительныйЛист
  - zup_2: Документ.НачислениеЗарплатыРаботникам (ТЧ Начисления + Удержания)

Запуск: C:\\Python313\\python.exe discovery_zup_struktura.py
"""
import io
import sys
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

ARTIFACT_DIR = os.path.join(os.path.dirname(__file__), "..", "_artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)

OUT_FILE = os.path.join(ARTIFACT_DIR, "zup_schema.md")


def dump_register(md_root, name, kind, lines):
    """kind — РегистрыНакопления | РегистрыРасчета"""
    try:
        if kind == "Накопл":
            md = md_root.РегистрыНакопления.Найти(name)
        else:
            md = md_root.РегистрыРасчета.Найти(name)
    except Exception as e:
        lines.append(f"\n### {kind}.{name} — ERROR {e}\n")
        return
    if md is None:
        lines.append(f"\n### {kind}.{name} — NOT FOUND\n")
        return
    lines.append(f"\n### {kind}.{name}\n")
    lines.append("**Измерения:** " + ", ".join([f"`{d.Имя}`" for d in md.Измерения]) + "\n\n")
    lines.append("**Ресурсы:** " + ", ".join([f"`{r.Имя}`" for r in md.Ресурсы]) + "\n\n")
    lines.append("**Реквизиты:** " + ", ".join([f"`{a.Имя}`" for a in md.Реквизиты]) + "\n\n")


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")

    lines = ["# Структуры регистров zup_1 + zup_2\n",
             "_Источник истины:_ `Documents/А_ОтражениеЗПпоКазне/Ext/ObjectModule.bsl` (linenums ~439 ZUP_1, ~828 zup_2)\n"]

    # ZUP_1 — регламентированный учёт зарплаты
    lines.append("\n## ZUP_1 (BAS ЗУП регламент)\n")
    try:
        zup1 = v8.Connect('Srvr="localhost";Ref="zup";Usr="cfo";Pwd="2442"')
        dump_register(zup1.Метаданные, "ВзаиморасчетыПоНДФЛ", "Накопл", lines)
        dump_register(zup1.Метаданные, "ВзносыВФонды", "Расчет", lines)
        dump_register(zup1.Метаданные, "УдержанияРаботниковОрганизаций", "Расчет", lines)

        # Тестовый запрос — посмотреть реальные данные ВзаиморасчетыПоНДФЛ за апрель 2026
        q = zup1.NewObject("Запрос")
        q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 3
    Регистр.Сотрудник.Физлицо.КодПоДРФО КАК ИНН,
    Регистр.Сотрудник.Физлицо.Наименование КАК ФИО,
    Регистр.ДоходНДФЛ.Код КАК КодДохода,
    Регистр.НалогПриход КАК НалогПриход,
    Регистр.ДоходПриход КАК ДоходПриход
ИЗ
    РегистрНакопления.ВзаиморасчетыПоНДФЛ.Обороты(&НачП, &КонП,,) КАК Регистр
ГДЕ Регистр.НалогПриход <> 0
"""
        q.УстановитьПараметр("НачП", zup1.NewObject("ДатаВремя", 2026, 4, 1))
        q.УстановитьПараметр("КонП", zup1.NewObject("ДатаВремя", 2026, 4, 30, 23, 59, 59))
        try:
            ts = q.Выполнить().Выгрузить()
            lines.append(f"\n**Тест ВзаиморасчетыПоНДФЛ за апр 2026:** найдено {ts.Количество()} строк\n")
            if ts.Количество() > 0:
                lines.append("| ИНН | ФИО | КодДохода | НалогПриход | ДоходПриход |\n")
                lines.append("|---|---|---|---:|---:|\n")
                for i in range(min(3, ts.Количество())):
                    r = ts.Получить(i)
                    lines.append(f"| {r.ИНН} | {r.ФИО} | {r.КодДохода} | {r.НалогПриход} | {r.ДоходПриход} |\n")
        except Exception as e:
            info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
            lines.append(f"\n**ERROR test query:** {info}\n")
    except Exception as e:
        lines.append(f"\n**ERROR connect zup_1:** {e}\n")

    # zup_2 — управленческий учёт
    lines.append("\n## zup_2 (BAS ЗУП управленческий)\n")
    try:
        zup2 = v8.Connect('Srvr="localhost";Ref="zup2";Usr="cfo";Pwd="2442"')
        try:
            doc_md = zup2.Метаданные.Документы.Найти("НачислениеЗарплатыРаботникам")
            if doc_md:
                lines.append("\n### Документ.НачислениеЗарплатыРаботникам\n")
                lines.append("**Реквизиты:** " + ", ".join([f"`{a.Имя}`" for a in doc_md.Реквизиты]) + "\n\n")
                lines.append("**Табличные части:**\n")
                for ts_md in doc_md.ТабличныеЧасти:
                    attrs = ", ".join([f"`{a.Имя}`" for a in ts_md.Реквизиты])
                    lines.append(f"- `{ts_md.Имя}`: {attrs}\n")
        except Exception as e:
            lines.append(f"\n**ERROR doc metadata:** {e}\n")

        # Тестовый запрос — посмотреть данные за апрель 2026
        q = zup2.NewObject("Запрос")
        q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 3
    Начисления.Физлицо.КодПоДРФО КАК ИНН,
    Начисления.Физлицо.Наименование КАК ФИО,
    Начисления.ВидРасчета.Наименование КАК ВидРасчета,
    Начисления.Результат КАК Сумма
ИЗ
    Документ.НачислениеЗарплатыРаботникам.Начисления КАК Начисления
ГДЕ
    Начисления.Ссылка.Дата МЕЖДУ &НачП И &КонП
    И Начисления.Ссылка.Проведен = ИСТИНА
    И Начисления.Результат <> 0
"""
        q.УстановитьПараметр("НачП", zup2.NewObject("ДатаВремя", 2026, 4, 1))
        q.УстановитьПараметр("КонП", zup2.NewObject("ДатаВремя", 2026, 4, 30, 23, 59, 59))
        try:
            ts = q.Выполнить().Выгрузить()
            lines.append(f"\n**Тест НачислениеЗарплатыРаботникам.Начисления за апр 2026:** найдено {ts.Количество()} строк\n")
            if ts.Количество() > 0:
                lines.append("| ИНН | ФИО | ВидРасчета | Сумма |\n")
                lines.append("|---|---|---|---:|\n")
                for i in range(min(3, ts.Количество())):
                    r = ts.Получить(i)
                    lines.append(f"| {r.ИНН} | {r.ФИО} | {r.ВидРасчета} | {r.Сумма} |\n")
        except Exception as e:
            info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
            lines.append(f"\n**ERROR test query:** {info}\n")
    except Exception as e:
        lines.append(f"\n**ERROR connect zup_2:** {e}\n")

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.writelines(lines)
    print(f"OK -> {OUT_FILE}")


if __name__ == "__main__":
    main()
