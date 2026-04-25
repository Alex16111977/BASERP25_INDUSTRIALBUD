"""
Додає реквізит А_НомерДокументаИзБух (Строка(11)) у 22 документи конфігурації
через цикл викликів meta-edit.ps1, потім нормалізує Synonym (ru+uk) через regex.

Запуск (один раз):
    C:\\Python313\\python.exe scripts/_add_nomer_iz_buh_to_docs.py

Pre-condition: 1С Конфігуратор закритий (інакше db-load-xml потім впаде).
"""
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(r"C:\Configuration_downloads\BASERP25")
META_EDIT = ROOT / ".claude" / "skills" / "meta-edit" / "scripts" / "meta-edit.ps1"

DOCS = [
    "ПриходныйКассовыйОрдер",
    "ВозвратТоваровПоставщику",
    "ЗаказКлиента",
    "ЗаказПоставщику",
    "ПеремещениеТоваров",
    "ПоступлениеБезналичныхДенежныхСредств",
    "СписаниеБезналичныхДенежныхСредств",
    "АвансовыйОтчет",
    "РасходныйКассовыйОрдер",
    "РеализацияТоваровУслуг",
    "ПриобретениеТоваровУслуг",
    "СборкаТоваров",
    "АктВыполненныхРабот",
    "ПокупкаПродажаВалюты",
    "ЗаявкаНаПокупкуПродажуВалюты",
    "ВнутреннееПотреблениеТоваров",
    "ПриобретениеУслугПрочихАктивов",
    "ВзаимозачетЗадолженности",
    "А_ОтражениеЗарплатыВУчете",
    "ПередачаМатериаловВПроизводство",
    "ПринятиеКУчетуОС",
    "ПринятиеКУчетуНМА",
]

ATTR_MARKER = "<Name>А_НомерДокументаИзБух</Name>"

SYNONYM_PATTERN = re.compile(
    r'(<Attribute uuid="[^"]+">\s*<Properties>\s*<Name>А_НомерДокументаИзБух</Name>\s*'
    r'<Synonym>)(.*?)(</Synonym>)',
    re.DOTALL,
)
NEW_SYNONYM_INNER = (
    "\n\t\t\t\t\t\t<v8:item>\n\t\t\t\t\t\t\t<v8:lang>ru</v8:lang>\n"
    "\t\t\t\t\t\t\t<v8:content>Номер документа из бух</v8:content>\n"
    "\t\t\t\t\t\t</v8:item>\n"
    "\t\t\t\t\t\t<v8:item>\n\t\t\t\t\t\t\t<v8:lang>uk</v8:lang>\n"
    "\t\t\t\t\t\t\t<v8:content>Номер вимога накладна з бух.</v8:content>\n"
    "\t\t\t\t\t\t</v8:item>\n\t\t\t\t\t"
)


def fix_synonym(xml_path: Path):
    text = xml_path.read_text(encoding="utf-8")
    new_text, n = SYNONYM_PATTERN.subn(
        lambda m: m.group(1) + NEW_SYNONYM_INNER + m.group(3),
        text,
        count=1,
    )
    if n == 0:
        return False
    if new_text != text:
        xml_path.write_text(new_text, encoding="utf-8")
        return True
    return False


def add_attribute(doc: str) -> tuple[bool, str]:
    """Returns (success, message)."""
    xml = ROOT / "Documents" / f"{doc}.xml"
    if not xml.exists():
        return False, f"NOFILE: {xml}"
    if ATTR_MARKER in xml.read_text(encoding="utf-8"):
        return True, "SKIP (already has)"

    cmd = [
        "powershell.exe", "-NoProfile", "-File", str(META_EDIT),
        "-ObjectPath", f"Documents/{doc}",
        "-Operation", "add-attribute",
        "-Value", "А_НомерДокументаИзБух: Строка(11)",
        "-NoValidate",
    ]
    res = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True, encoding="utf-8")
    if res.returncode != 0:
        return False, f"meta-edit failed: {res.stderr.strip()[:300]}"
    if ATTR_MARKER not in xml.read_text(encoding="utf-8"):
        return False, "Attribute not in XML after meta-edit"
    fixed = fix_synonym(xml)
    return True, f"ADDED + Synonym {'fixed' if fixed else 'NOT fixed'}"


def main():
    print(f"Documents to process: {len(DOCS)}")
    print()
    results = []
    for i, doc in enumerate(DOCS, 1):
        ok, msg = add_attribute(doc)
        marker = "OK " if ok else "FAIL"
        print(f"[{i:>2}/{len(DOCS)}] {marker}  {doc:<40} {msg}")
        results.append((doc, ok, msg))

    ok_count = sum(1 for _, o, _ in results if o)
    print()
    print(f"Підсумок: {ok_count}/{len(DOCS)} OK")
    if ok_count < len(DOCS):
        sys.exit(1)


if __name__ == "__main__":
    main()
