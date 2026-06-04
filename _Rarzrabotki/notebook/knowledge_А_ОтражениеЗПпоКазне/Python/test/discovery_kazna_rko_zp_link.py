"""
Discovery: связь РКО Казны с РаспределениеЗП — как Казна оформляет выплату ЗП.

Гипотезы:
- РКО.ДокументОснование → ссылка на РаспределениеЗП?
- РКО.Расшифровка.Подразделение → совпадает с Распределение.Подразделение?
- ВидОперации/ХозяйственнаяОперация — какой признак указывает на «выплата ЗП»?

Запуск: C:\\Python313\\python.exe discovery_kazna_rko_zp_link.py
"""
import io
import sys
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

ARTIFACT_DIR = os.path.join(os.path.dirname(__file__), "..", "_artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)
OUT_FILE = os.path.join(ARTIFACT_DIR, "kazna_rko_zp_link.md")


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    kazna = v8.Connect('Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"')

    lines = ["# Связь РКО Казны с РаспределениеЗП (выплата ЗП по подразделениям)\n",
             "_Discovery 2026-05-26_\n"]

    # 1. Тип реквизита ДокументОснование в РКО
    lines.append("\n## 1. Реквизит РКО.ДокументОснование — какие типы документов поддерживает\n")
    try:
        md = kazna.Метаданные.Документы.Найти("РасходныйКассовыйОрдер")
        do_attr = md.Реквизиты.Найти("ДокументОснование")
        if do_attr is not None:
            # ТипЗначения — composite type
            tip = do_attr.Тип
            types_list = []
            for t in tip.Типы():
                try:
                    types_list.append(str(t))
                except:
                    pass
            lines.append(f"Типы (count={len(types_list)}):\n")
            for t in types_list:
                if "Документ" in t or "Распределение" in t:
                    lines.append(f"- `{t}`\n")
    except Exception as e:
        lines.append(f"**ERROR meta:** {e}\n")

    # 2. ВидОперации в РКО — найти связанное с ЗП
    lines.append("\n## 2. РКО.ВидОперации — какие виды связаны с ЗП\n")
    try:
        vo_attr = md.Реквизиты.Найти("ВидОперации")
        if vo_attr is not None:
            tip = vo_attr.Тип
            for t in tip.Типы():
                t_str = str(t)
                lines.append(f"- тип: `{t_str}`\n")
                # Если это перечисление — попробовать вывести значения
                if "Перечисление" in t_str:
                    enum_name = t_str.split(".")[-1] if "." in t_str else None
                    if enum_name:
                        try:
                            enum_md = kazna.Метаданные.Перечисления.Найти(enum_name)
                            if enum_md is not None:
                                lines.append("  Значения:\n")
                                for v in enum_md.ЗначенияПеречисления:
                                    lines.append(f"    - `{v.Имя}`\n")
                        except Exception as ee:
                            lines.append(f"  ERR enum: {ee}\n")
    except Exception as e:
        lines.append(f"**ERROR vid op:** {e}\n")

    # 3. Найти РКО с ДокументОснование = РаспределениеЗП (любой проведённый)
    lines.append("\n## 3. Образец РКО Казны с основанием РаспределениеЗП (если есть)\n")
    q = kazna.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 5
    РКО.Ссылка КАК Ссылка,
    РКО.Номер  КАК Номер,
    РКО.Дата   КАК Дата,
    РКО.СуммаДокумента КАК Сумма,
    РКО.ДокументОснование КАК ДокОсн,
    РКО.Контрагент.Наименование КАК КонтрагентИмя,
    РКО.Подразделение.Код КАК ПодрКод,
    РКО.СтатьяДвиженияДенежныхСредств.Код КАК СтатьяКод,
    РКО.ВидОперации КАК ВидОп
ИЗ Документ.РасходныйКассовыйОрдер КАК РКО
ГДЕ РКО.Проведен = ИСТИНА
    И РКО.ДокументОснование ССЫЛКА Документ.РаспределениеЗаработнойПлаты
УПОРЯДОЧИТЬ ПО РКО.Дата УБЫВ
"""
    try:
        rs = q.Выполнить()
        if rs.Пустой():
            lines.append("_Нет РКО с ДокументОснование=РаспределениеЗП в Казне._\n")
        else:
            sel = rs.Выбрать()
            lines.append("| Номер | Дата | Сумма | КонтрагентИмя | ПодрКод | СтатьяКод | ВидОп | ДокОснование |\n")
            lines.append("|---|---|---:|---|---|---|---|---|\n")
            while sel.Следующий():
                lines.append(f"| {sel.Номер} | {sel.Дата} | {sel.Сумма} | {sel.КонтрагентИмя} | "
                             f"{sel.ПодрКод} | {sel.СтатьяКод} | {sel.ВидОп} | {sel.ДокОсн} |\n")
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        lines.append(f"**ERROR rko query:** {info}\n")

    # 4. Регистр движений по выплате ЗП — поискать что пишет в БДДС РКО
    lines.append("\n## 4. РКО Казны: движения в РегНакопл.БДДС (последний за май 2026)\n")
    q = kazna.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1
    РКО.Ссылка КАК Ссылка
ИЗ Документ.РасходныйКассовыйОрдер КАК РКО
ГДЕ РКО.Проведен = ИСТИНА
    И ГОД(РКО.Дата) = 2026 И МЕСЯЦ(РКО.Дата) = 5
УПОРЯДОЧИТЬ ПО РКО.Дата УБЫВ
"""
    try:
        rs = q.Выполнить()
        if not rs.Пустой():
            sel = rs.Выбрать(); sel.Следующий()
            rko_ref = sel.Ссылка
            lines.append(f"- Тест.документ: `{rko_ref}`\n")

            # Движения регистров
            obj = rko_ref.ПолучитьОбъект()
            md = kazna.Метаданные.Документы.Найти("РасходныйКассовыйОрдер")
            # Список движений
            for r_md in md.Движения:
                try:
                    movements = obj.Движения.Найти(r_md.Имя)
                    if movements is not None:
                        movements.Прочитать()
                        cnt = movements.Количество()
                        lines.append(f"- РегНакопл `{r_md.Имя}`: {cnt} строк\n")
                except Exception as ee:
                    lines.append(f"  ERR {r_md.Имя}: {ee}\n")
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        lines.append(f"**ERROR rko movement:** {info}\n")

    # 5. БДДС: есть ли строки, где Регистратор = РасходныйКассовыйОрдер и сумма по ЗП-статьям?
    lines.append("\n## 5. БДДС с Регистратор=РКО за май 2026 (по сотрудникам)\n")
    q = kazna.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 5
    БДДС.Регистратор КАК Регистратор,
    БДДС.Сотрудник.Наименование КАК СотрудникФИО,
    БДДС.Сотрудник.ИНН КАК СотрудникИНН,
    БДДС.Подразделение.Код КАК ПодрКод,
    БДДС.СтатьяДвиженияДенежныхСредств.Код КАК СтатьяКод,
    БДДС.Сумма КАК Сумма,
    БДДС.СуммаНачисления КАК СуммаНачисления,
    БДДС.СуммаФОТ КАК СуммаФОТ,
    БДДС.СуммаНалогов КАК СуммаНалогов
ИЗ РегистрНакопления.БДДС КАК БДДС
ГДЕ БДДС.Регистратор ССЫЛКА Документ.РасходныйКассовыйОрдер
    И ГОД(БДДС.Период) = 2026 И МЕСЯЦ(БДДС.Период) = 5
    И БДДС.Сотрудник <> ЗНАЧЕНИЕ(Справочник.Сотрудники.ПустаяСсылка)
"""
    try:
        rs = q.Выполнить()
        if rs.Пустой():
            lines.append("_Нет БДДС с РКО+Сотрудник за май 2026 — возможно РКО пишет БДДС без сотрудника._\n")
        else:
            sel = rs.Выбрать()
            lines.append("| Рег | ФИО | ИНН | ПодрКод | СтатьяКод | Сумма | НачКаз | ФОТ | Нал |\n")
            lines.append("|---|---|---|---|---|---:|---:|---:|---:|\n")
            while sel.Следующий():
                lines.append(f"| {sel.Регистратор} | {sel.СотрудникФИО} | {sel.СотрудникИНН} | "
                             f"{sel.ПодрКод} | {sel.СтатьяКод} | {sel.Сумма} | {sel.СуммаНачисления} | "
                             f"{sel.СуммаФОТ} | {sel.СуммаНалогов} |\n")
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        lines.append(f"**ERROR bdds rko:** {info}\n")

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.writelines(lines)
    print(f"OK -> {OUT_FILE}")


if __name__ == "__main__":
    main()
