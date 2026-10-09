# -*- coding: utf-8 -*-
"""Метадані «обсяг за договором» (09.10.2026, специфікація 2026-10-09-finzvit-obyem-metal-domiki-design):
регістр А_ПараметрыДоговоровФинотчета +ВидПродукции (Строка 20) +ОбъемПоДоговору (Число 15,3 ≥ 0);
ТЧ Расчет +ОбъемПоДоговору +ОбъемОтгружено +ОбъемПодписано (Число 15,3) +ЕдиницаОбъема (Строка 5);
ТЧ РасшифровкаДокументов +Количество (Число 15,3).
Накладається на ВИВАНТАЖЕННЯ З БАЗИ (worktree синхронізовано комітом 829d2c2dec). Ідемпотентний; minidom + assert.
Запуск: python doc_finrep_meta_0910_obyem.py <корінь BASEBuh>"""
import io
import re
import sys
import uuid
from xml.dom import minidom

sys.stdout.reconfigure(encoding="utf-8")
КОРІНЬ = sys.argv[1]
РЕГ = КОРІНЬ + r"\InformationRegisters\А_ПараметрыДоговоровФинотчета.xml"
ДОК = КОРІНЬ + r"\Documents\А_ФинансовыйОтчетПроизводства.xml"


def читати(п):
    return io.open(п, encoding="utf-8-sig").read()


def писати(п, т):
    io.open(п, "w", encoding="utf-8-sig", newline="").write(т)


def тип_число(digits, frac, знак, в):
    return (f"{в}\t<v8:Type>xs:decimal</v8:Type>\n{в}\t<v8:NumberQualifiers>\n"
            f"{в}\t\t<v8:Digits>{digits}</v8:Digits>\n{в}\t\t<v8:FractionDigits>{frac}</v8:FractionDigits>\n"
            f"{в}\t\t<v8:AllowedSign>{знак}</v8:AllowedSign>\n{в}\t</v8:NumberQualifiers>")


def тип_рядок(довжина, в):
    return (f"{в}\t<v8:Type>xs:string</v8:Type>\n{в}\t<v8:StringQualifiers>\n"
            f"{в}\t\t<v8:Length>{довжина}</v8:Length>\n{в}\t\t<v8:AllowedLength>Variable</v8:AllowedLength>\n"
            f"{в}\t</v8:StringQualifiers>")


def клон(д, тег, зразок, імя, ru, uk, тип):
    """Копія <тег uuid=…> з іменем-зразком (унікальним у файлі) під новим ім'ям одразу після нього.
    тип — функція відступу → вміст <Type>."""
    if f"<Name>{імя}</Name>" in д:
        return д
    assert д.count(f"<Name>{зразок}</Name>") == 1, зразок
    м = re.search(f"<Name>{зразок}</Name>", д)
    початок = д.rfind(f"<{тег} uuid=", 0, м.start())
    кінець = д.find(f"</{тег}>", м.end()) + len(f"</{тег}>")
    новий = re.sub(f'<{тег} uuid="[^"]+">', f'<{тег} uuid="{uuid.uuid4()}">', д[початок:кінець], count=1)
    новий = новий.replace(f"<Name>{зразок}</Name>", f"<Name>{імя}</Name>")
    новий = re.sub(r"(<v8:lang>ru</v8:lang>\s*<v8:content>)[^<]*", lambda x: x.group(1) + ru, новий, count=1)
    новий = re.sub(r"(<v8:lang>uk</v8:lang>\s*<v8:content>)[^<]*", lambda x: x.group(1) + uk, новий, count=1)
    т = re.search(r"(?P<в>[ \t]*)<Type>[\s\S]*?</Type>", новий)
    в = т.group("в")
    новий = новий[:т.start()] + f"{в}<Type>\n{тип(в)}\n{в}</Type>" + новий[т.end():]
    # ToolTip/Format зразка не переносимо — порожні
    новий = re.sub(r"<Format>[\s\S]*?</Format>", "<Format/>", новий)
    новий = re.sub(r"<ToolTip>[\s\S]*?</ToolTip>", "<ToolTip/>", новий)
    відступ = д[д.rfind("\n", 0, початок) + 1:початок]
    return д[:кінець] + "\n" + відступ + новий + д[кінець:]


# --- регістр: клон ресурсів (зразки — ПравилоНДС рядок, СуммаПредоплаты число)
р = читати(РЕГ)
assert "\r\n" not in р
р = клон(р, "Resource", "ПравилоНДС", "ВидПродукции", "Вид продукции (металлоконструкции / дома)",
         "Вид продукції (металоконструкції / будинки)", lambda в: тип_рядок(20, в))
р = клон(р, "Resource", "СуммаПредоплаты", "ОбъемПоДоговору", "Объём по договору (т / дома)",
         "Обсяг за договором (т / будинки)", lambda в: тип_число(15, 3, "Nonnegative", в))
писати(РЕГ, р)
minidom.parse(РЕГ)
р = читати(РЕГ)
assert р.count("<Resource uuid=") == 18, р.count("<Resource uuid=")

# --- документ: ТЧ Расчет (після ОстатокАванса) і РасшифровкаДокументов (після СуммаКт)
д = читати(ДОК)
assert "\r\n" not in д
ч153 = lambda в: тип_число(15, 3, "Any", в)
д = клон(д, "Attribute", "ОстатокАванса", "ОбъемПоДоговору", "Объём по договору", "Обсяг за договором", ч153)
д = клон(д, "Attribute", "ОбъемПоДоговору", "ОбъемОтгружено", "Отгружено (акты)", "Відвантажено (акти)", ч153)
д = клон(д, "Attribute", "ОбъемОтгружено", "ОбъемПодписано", "в т.ч. подписано", "з них підписано", ч153)
д = клон(д, "Attribute", "ОбъемПодписано", "ЕдиницаОбъема", "Ед. объёма", "Од. обсягу", lambda в: тип_рядок(5, в))
д = клон(д, "Attribute", "СуммаКт", "Количество", "Количество (т / дома)", "Кількість (т / будинки)", ч153)
писати(ДОК, д)
minidom.parse(ДОК)
д = читати(ДОК)
for імя in ("ОбъемПоДоговору", "ОбъемОтгружено", "ОбъемПодписано", "ЕдиницаОбъема", "Количество"):
    assert д.count(f"<Name>{імя}</Name>") == 1, імя
print("OK: регістр 18 ресурсів; Расчет +4; РасшифровкаДокументов +Количество")
