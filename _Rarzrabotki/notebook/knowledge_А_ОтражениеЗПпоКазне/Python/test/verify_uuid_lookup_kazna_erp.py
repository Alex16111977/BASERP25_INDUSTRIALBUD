"""
Verify: cross-base UUID lookup Казна↔ERP.

Архитектура: правила обмена _Rarzrabotki/ConvertERP/ переносят документы
с UUID 1:1, поэтому ERP-документ можно найти НАПРЯМУЮ по UUID регистратора
из Казни (без РегСведений.СоответствияОбъектовИнформационныхБаз и без
поиска по Дата+Номер).

Проверяемые типы:
  - Документ.РаспределениеЗаработнойПлаты (Казна) → Документ.А_РаспределениеЗаработнойПлаты (ERP)
  - Документ.РаспределениеФ2 (Казна) → Документ.РаспределениеФ2 (ERP)

Паттерн: kazna.string(doc.Ссылка.УникальныйИдентификатор()) — английский lowercase
(см. CLAUDE.md «Cross-base UUID lookup»).

Запуск: C:\\Python313\\python.exe verify_uuid_lookup_kazna_erp.py
"""
import io
import sys
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

ARTIFACT_DIR = os.path.join(os.path.dirname(__file__), "..", "_artifacts")
os.makedirs(ARTIFACT_DIR, exist_ok=True)

OUT_FILE = os.path.join(ARTIFACT_DIR, "uuid_pair_example.md")


def find_one(kazna, doc_type_name, lines):
    """Найти один документ заданного типа в Казне и проверить UUID lookup в ERP."""
    q = kazna.NewObject("Запрос")
    q.Текст = f"""
ВЫБРАТЬ ПЕРВЫЕ 1
    Документ.Ссылка КАК Ссылка,
    Документ.Номер КАК Номер,
    Документ.Дата КАК Дата
ИЗ
    Документ.{doc_type_name} КАК Документ
ГДЕ Документ.Проведен = ИСТИНА
УПОРЯДОЧИТЬ ПО Документ.Дата УБЫВ
"""
    try:
        rs = q.Выполнить()
        if rs.Пустой():
            lines.append(f"\n**{doc_type_name}** — нет проведённых документов в Казне.\n")
            return None, None
        sel = rs.Выбрать()
        sel.Следующий()
        uid_str = kazna.string(sel.Ссылка.УникальныйИдентификатор())
        lines.append(f"\n### {doc_type_name} (Казна)\n")
        lines.append(f"- Номер: `{sel.Номер}`\n")
        lines.append(f"- Дата:  `{sel.Дата}`\n")
        lines.append(f"- UUID:  `{uid_str}`\n")
        return uid_str, (sel.Номер, sel.Дата)
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        lines.append(f"\n**{doc_type_name}** — ERROR find: {info}\n")
        return None, None


def check_erp(erp, erp_type, uid_str, lines):
    """Проверить наличие документа в ERP по UUID."""
    if not uid_str:
        return
    try:
        uid = erp.NewObject("УникальныйИдентификатор", uid_str)
        # Доступ к коллекции документов через атрибут .Документы
        col = getattr(erp.Документы, erp_type)
        ref = col.ПолучитьСсылку(uid)
        obj = ref.ПолучитьОбъект()
        if obj is None:
            lines.append(f"\n**ERP.{erp_type}** — UUID не маппится (broken reference)\n")
        else:
            lines.append(f"\n**ERP.{erp_type}** ✓ найден\n")
            lines.append(f"- Номер: `{obj.Номер}`\n")
            lines.append(f"- Дата:  `{obj.Дата}`\n")
            lines.append(f"- UUID совпал 1:1 (правила обмена)\n")
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        lines.append(f"\n**ERP.{erp_type}** — ERROR: {info}\n")


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    kazna = v8.Connect('Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"')
    erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

    lines = ["# UUID lookup Казна → ERP (1:1)\n",
             "Архитектура: правила обмена ConvertERP/Казна сохраняют UUID при репликации документов.\n",
             "Паттерн: `kazna.string(doc.Ссылка.УникальныйИдентификатор())` (английский lowercase).\n"]

    # Пара 1: РаспределениеЗаработнойПлаты
    lines.append("\n## Пара #1: РаспределениеЗП\n")
    uid1, _ = find_one(kazna, "РаспределениеЗаработнойПлаты", lines)
    check_erp(erp, "А_РаспределениеЗаработнойПлаты", uid1, lines)

    # Пара 2: РаспределениеФ2
    lines.append("\n## Пара #2: РаспределениеФ2\n")
    uid2, _ = find_one(kazna, "РаспределениеФ2", lines)
    check_erp(erp, "РаспределениеФ2", uid2, lines)

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.writelines(lines)
    print(f"OK -> {OUT_FILE}")


if __name__ == "__main__":
    main()
