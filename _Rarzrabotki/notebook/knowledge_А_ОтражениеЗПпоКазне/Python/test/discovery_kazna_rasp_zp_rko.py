"""
Discovery: структура документов Казны, связанных с выплатой ЗП:
- Документ.РаспределениеЗаработнойПлаты (Казна) — реквизиты + ТЧ
- Документ.РаспределениеФ2 (Казна) — реквизиты + ТЧ
- Документ.РасходныйКассовыйОрдер (Казна) — выплаты ЗП

Цель: понять, как Казна формирует РКО на выплату ЗП и как они связаны с РаспределениеЗП/Ф2.

Запуск: C:\\Python313\\python.exe discovery_kazna_rasp_zp_rko.py
"""
import io
import sys
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

ARTIFACT_DIR = os.path.join(os.path.dirname(__file__), "..", "_artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)
OUT_FILE = os.path.join(ARTIFACT_DIR, "kazna_rasp_zp_rko_schema.md")


def dump_doc(meta_root, name, lines):
    """Документировать реквизиты + ТЧ документа."""
    try:
        md = meta_root.Документы.Найти(name)
        if md is None:
            lines.append(f"\n### Документ.{name} — NOT FOUND\n")
            return None
    except Exception as e:
        lines.append(f"\n### Документ.{name} — ERROR {e}\n")
        return None

    lines.append(f"\n### Документ.{name}\n")
    lines.append("**Реквизиты:**\n")
    for a in md.Реквизиты:
        lines.append(f"- `{a.Имя}` — {a.Тип}\n")
    lines.append("\n**Табличные части:**\n")
    for tc in md.ТабличныеЧасти:
        attrs = ", ".join([f"`{a.Имя}`" for a in tc.Реквизиты])
        lines.append(f"- `{tc.Имя}`: {attrs}\n")
    return md


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    kazna = v8.Connect('Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"')

    lines = ["# Структура документов Казны: РаспределениеЗП / РаспределениеФ2 / РКО\n",
             "_Discovery 2026-05-26_\n"]

    # 1. РаспределениеЗаработнойПлаты
    lines.append("\n## 1. Документ.РаспределениеЗаработнойПлаты (Казна)\n")
    dump_doc(kazna.Метаданные, "РаспределениеЗаработнойПлаты", lines)

    # 2. РаспределениеФ2
    lines.append("\n## 2. Документ.РаспределениеФ2 (Казна)\n")
    dump_doc(kazna.Метаданные, "РаспределениеФ2", lines)

    # 3. РасходныйКассовыйОрдер
    lines.append("\n## 3. Документ.РасходныйКассовыйОрдер (Казна)\n")
    dump_doc(kazna.Метаданные, "РасходныйКассовыйОрдер", lines)

    # 4. Тестовые данные — 1 последний РаспределениеЗП с шапкой и ТЧ
    lines.append("\n## 4. Образец РаспределениеЗП (последний проведённый)\n")
    q = kazna.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1
    Док.Ссылка КАК Ссылка,
    Док.Номер  КАК Номер,
    Док.Дата   КАК Дата,
    Док.Организация КАК Организация,
    Док.КассаДатаС КАК КассаДатаС,
    Док.КассаДатаПо КАК КассаДатаПо
ИЗ Документ.РаспределениеЗаработнойПлаты КАК Док
ГДЕ Док.Проведен = ИСТИНА
УПОРЯДОЧИТЬ ПО Док.Дата УБЫВ
"""
    try:
        rs = q.Выполнить()
        if not rs.Пустой():
            sel = rs.Выбрать(); sel.Следующий()
            lines.append(f"- Ссылка: `{sel.Ссылка}`\n")
            lines.append(f"- Номер:  `{sel.Номер}`\n")
            lines.append(f"- Дата:   `{sel.Дата}`\n")
            lines.append(f"- КассаДатаС/По: `{sel.КассаДатаС}` .. `{sel.КассаДатаПо}`\n")
            lines.append(f"- UUID: `{kazna.string(sel.Ссылка.УникальныйИдентификатор())}`\n")

            obj = sel.Ссылка.ПолучитьОбъект()
            if obj is not None:
                # Перечислить ТЧ и их размеры
                md = kazna.Метаданные.Документы.Найти("РаспределениеЗаработнойПлаты")
                for tc_md in md.ТабличныеЧасти:
                    tc = getattr(obj, tc_md.Имя)
                    cnt = tc.Количество()
                    lines.append(f"\n**ТЧ `{tc_md.Имя}` — {cnt} строк**\n")
                    if cnt > 0:
                        attr_names = [a.Имя for a in tc_md.Реквизиты][:7]
                        lines.append("| " + " | ".join(attr_names) + " |\n")
                        lines.append("|" + "|".join(["---"] * len(attr_names)) + "|\n")
                        for i in range(min(3, cnt)):
                            r = tc.Получить(i)
                            vals = []
                            for n in attr_names:
                                v = getattr(r, n)
                                s = str(v)
                                if len(s) > 30:
                                    s = s[:27] + "..."
                                vals.append(s)
                            lines.append("| " + " | ".join(vals) + " |\n")
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        lines.append(f"\n**ERROR sample query:** {info}\n")

    # 5. РКО на выплату ЗП — поискать через ХозяйственнаяОперация
    lines.append("\n## 5. РКО Казны на выплату ЗП (последние 5 за май 2026)\n")
    q = kazna.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 5
    РКО.Ссылка КАК Ссылка,
    РКО.Номер  КАК Номер,
    РКО.Дата   КАК Дата,
    РКО.ХозяйственнаяОперация КАК ХозОп,
    РКО.СуммаДокумента КАК Сумма
ИЗ Документ.РасходныйКассовыйОрдер КАК РКО
ГДЕ РКО.Проведен = ИСТИНА
    И РКО.ХозяйственнаяОперация ССЫЛКА Перечисление.ХозяйственныеОперации
УПОРЯДОЧИТЬ ПО РКО.Дата УБЫВ
"""
    try:
        rs = q.Выполнить()
        sel = rs.Выбрать()
        cnt = 0
        lines.append("| Номер | Дата | ХозОп | Сумма |\n")
        lines.append("|---|---|---|---:|\n")
        while sel.Следующий() and cnt < 5:
            lines.append(f"| {sel.Номер} | {sel.Дата} | {sel.ХозОп} | {sel.Сумма} |\n")
            cnt += 1
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        lines.append(f"\n**ERROR rko query:** {info}\n")

    # 6. ХозОп для выплаты ЗП — найти перечисление
    lines.append("\n## 6. Перечисление.ХозяйственныеОперации — значения связанные с ЗП\n")
    try:
        enum_md = kazna.Метаданные.Перечисления.Найти("ХозяйственныеОперации")
        if enum_md is not None:
            zp_related = []
            for v in enum_md.ЗначенияПеречисления:
                name = v.Имя
                if "Зарплат" in name or "ЗП" in name or "Выплат" in name or "Аванс" in name:
                    zp_related.append(name)
            lines.append("**Найденные значения с ЗП/Зарплата/Выплата/Аванс:**\n")
            for n in zp_related:
                lines.append(f"- `{n}`\n")
    except Exception as e:
        lines.append(f"\n**ERROR enum dump:** {e}\n")

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.writelines(lines)
    print(f"OK -> {OUT_FILE}")


if __name__ == "__main__":
    main()
