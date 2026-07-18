"""
Verify: один реальный документ А_ОтражениеЗПпоКазне — содержимое 7 ТЧ.

Цель: эталонный snapshot тестового документа для FINDINGS.md
(чтобы при будущих регрессиях знать ожидаемое содержимое).

Запуск: C:\\Python313\\python.exe verify_test_doc_loading.py
"""
import io
import sys
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

ARTIFACT_DIR = os.path.join(os.path.dirname(__file__), "..", "_artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)

OUT_FILE = os.path.join(ARTIFACT_DIR, "test_doc_sample.md")

TC_NAMES = [
    "РаспределениеКазна",
    "НачисленнаяЗарплатаИВзносыПоФизлицам",
    "НачисленныйНДФЛ",
    "НалогиБухгалтерия",
    "УдержанияБухгалтерия",
    "НачисленияУпр",
    "УдержанияУпр",
]


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

    # Найти один документ А_ОтражениеЗПпоКазне (любой, самый свежий)
    q = erp.NewObject("Запрос")
    q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1
    Док.Ссылка КАК Ссылка,
    Док.Номер  КАК Номер,
    Док.Дата   КАК Дата,
    Док.Организация.Наименование КАК Орг,
    Док.Проведен КАК Проведен
ИЗ
    Документ.А_ОтражениеЗПпоКазне КАК Док
УПОРЯДОЧИТЬ ПО Док.Дата УБЫВ
"""
    rs = q.Выполнить()
    if rs.Пустой():
        with open(OUT_FILE, "w", encoding="utf-8") as f:
            f.write("# А_ОтражениеЗПпоКазне — нет ни одного документа\n")
        print(f"OK (empty) -> {OUT_FILE}")
        return

    sel = rs.Выбрать()
    sel.Следующий()

    lines = [
        "# А_ОтражениеЗПпоКазне — образец документа\n",
        f"\n**Номер:** `{sel.Номер}`\n",
        f"\n**Дата:** `{sel.Дата}`\n",
        f"\n**Организация:** {sel.Орг}\n",
        f"\n**Проведен:** {sel.Проведен}\n",
        f"\n**UUID:** `{erp.string(sel.Ссылка.УникальныйИдентификатор())}`\n",
    ]

    obj = sel.Ссылка.ПолучитьОбъект()
    if obj is None:
        lines.append("\n**ERROR:** документ не найден через ПолучитьОбъект()\n")
        with open(OUT_FILE, "w", encoding="utf-8") as f:
            f.writelines(lines)
        return

    # Подсчёт строк в каждой ТЧ + первые 3 строки
    for tc_name in TC_NAMES:
        try:
            tc = getattr(obj, tc_name)
            cnt = tc.Количество()
            lines.append(f"\n## ТЧ {tc_name} — {cnt} строк\n")
            if cnt > 0:
                # Динамически получить колонки из первой строки
                first = tc.Получить(0)
                # Имена реквизитов через метаданные
                md = erp.Метаданные.Документы.А_ОтражениеЗПпоКазне.ТабличныеЧасти.Найти(tc_name)
                if md is None:
                    lines.append(f"_(метаданные ТЧ не найдены)_\n")
                    continue
                attr_names = [a.Имя for a in md.Реквизиты]
                # Только первые 6 колонок (чтобы таблица была читабельной)
                show_names = attr_names[:6]
                lines.append("| " + " | ".join(show_names) + " |\n")
                lines.append("|" + "|".join(["---"] * len(show_names)) + "|\n")
                for i in range(min(3, cnt)):
                    r = tc.Получить(i)
                    vals = []
                    for n in show_names:
                        v = getattr(r, n)
                        s = str(v)
                        if len(s) > 30:
                            s = s[:27] + "..."
                        vals.append(s)
                    lines.append("| " + " | ".join(vals) + " |\n")
        except Exception as e:
            info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
            lines.append(f"\n## ТЧ {tc_name} — ERROR: {info}\n")

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.writelines(lines)
    print(f"OK -> {OUT_FILE}")


if __name__ == "__main__":
    main()
