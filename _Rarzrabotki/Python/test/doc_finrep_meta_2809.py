# -*- coding: utf-8 -*-
"""Метадані доробки 28.09.2026: 5 ресурсів регістру + ПодразделениеКазны у ТЧ РасшифровкаДокументов.
Ідемпотентний: повторний запуск нічого не змінює. Після запису — minidom і assert."""
import re, sys, uuid, io
from xml.dom import minidom
sys.stdout.reconfigure(encoding="utf-8")
КОРІНЬ = sys.argv[1] if len(sys.argv) > 1 else (
    r"C:\Configuration_downloads\BASERP25\.claude\worktrees\director-report-letter-689b66\_Rarzrabotki\BASEBuh")
РЕГ = КОРІНЬ + r"\InformationRegisters\А_ПараметрыДоговоровФинотчета.xml"
ДОК = КОРІНЬ + r"\Documents\А_ФинансовыйОтчетПроизводства.xml"

ХВІСТ = """				<PasswordMode>false</PasswordMode>
				<Format/>
				<EditFormat/>
				<ToolTip/>
				<MarkNegatives>false</MarkNegatives>
				<Mask/>
				<MultiLine>false</MultiLine>
				<ExtendedEdit>false</ExtendedEdit>
				<MinValue xsi:nil="true"/>
				<MaxValue xsi:nil="true"/>
				<FillFromFillingValue>false</FillFromFillingValue>
				<FillValue xsi:nil="true"/>
				<FillChecking>DontCheck</FillChecking>
				<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
				<ChoiceParameterLinks/>
				<ChoiceParameters/>
				<QuickChoice>Auto</QuickChoice>
				<CreateOnInput>Auto</CreateOnInput>
				<ChoiceForm/>
				<LinkByType/>
				<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
				<Indexing>DontIndex</Indexing>
				<FullTextSearch>Use</FullTextSearch>
				<DataHistory>Use</DataHistory>
			</Properties>
		</Resource>
"""
ЧИСЛО = """<v8:Type>xs:decimal</v8:Type>
					<v8:NumberQualifiers>
						<v8:Digits>15</v8:Digits>
						<v8:FractionDigits>2</v8:FractionDigits>
						<v8:AllowedSign>Any</v8:AllowedSign>
					</v8:NumberQualifiers>"""
ДАТА = """<v8:Type>xs:dateTime</v8:Type>
					<v8:DateQualifiers>
						<v8:DateFractions>Date</v8:DateFractions>
					</v8:DateQualifiers>"""
РЯДОК20 = """<v8:Type>xs:string</v8:Type>
					<v8:StringQualifiers>
						<v8:Length>20</v8:Length>
						<v8:AllowedLength>Variable</v8:AllowedLength>
					</v8:StringQualifiers>"""
РЕСУРСИ = [
    ("СуммаПредоплаты", "Сумма предоплаты", "Сума передоплати (аванс)", ЧИСЛО),
    ("ДатаЗачетаПредоплаты", "Дата зачета предоплаты", "Дата заліку передоплати", ДАТА),
    ("ОстатокПредоплатыНаДату", "Остаток предоплаты на дату", "Залишок передоплати на дату", ЧИСЛО),
    ("ПравилоНДС", "Правило НДС", "Правило ПДВ", РЯДОК20),
    ("СуммаГарантийногоУдержания", "Сумма гарантийного удержания", "Сума гарантійного утримання", ЧИСЛО),
]


def ресурс(ім, ru, uk, тип):
    return "".join("\t" + рядок if рядок.strip() else рядок
                   for рядок in _ресурс(ім, ru, uk, тип).splitlines(keepends=True))


def _ресурс(ім, ru, uk, тип):
    return f"""		<Resource uuid="{uuid.uuid4()}">
			<Properties>
				<Name>{ім}</Name>
				<Synonym>
					<v8:item>
						<v8:lang>ru</v8:lang>
						<v8:content>{ru}</v8:content>
					</v8:item>
					<v8:item>
						<v8:lang>uk</v8:lang>
						<v8:content>{uk}</v8:content>
					</v8:item>
				</Synonym>
				<Comment/>
				<Type>
					{тип}
				</Type>
""" + ХВІСТ


def читати(п):
    return io.open(п, encoding="utf-8-sig").read()


def писати(п, т):
    io.open(п, "w", encoding="utf-8-sig", newline="").write(т)


# --- регістр
т = читати(РЕГ)
перенос = "\r\n" if "\r\n" in т else "\n"
нові = "".join(ресурс(*р) for р in РЕСУРСИ if f"<Name>{р[0]}</Name>" not in т)
if нові:
    якір = "\t\t\t<Dimension uuid="
    assert т.count(якір) == 1, "якір Dimension"
    т = т.replace(якір, нові.replace("\n", перенос) + якір, 1)
т = т.replace("<v8:content>Процент гарантийного удержания</v8:content>",
              "<v8:content>Процент удержания / зачета аванса</v8:content>")
т = т.replace("<v8:content>Відсоток гарантійного утримання, %</v8:content>",
              "<v8:content>Відсоток утримання / заліку авансу, %</v8:content>")
писати(РЕГ, т)
minidom.parse(РЕГ)
т = читати(РЕГ)
for р in РЕСУРСИ:
    assert т.count(f"<Name>{р[0]}</Name>") == 1, р[0]
assert т.count("<Resource uuid=") == 10, т.count("<Resource uuid=")

# --- ТЧ РасшифровкаДокументов
д = читати(ДОК)
if "<Name>ПодразделениеКазны</Name>" not in д:
    м = re.search(r"<Name>СуммаКт</Name>", д)
    assert м and д.count("<Name>СуммаКт</Name>") == 1, "СуммаКт у ТЧ"
    початок = д.rfind("<Attribute uuid=", 0, м.start())
    кінець = д.find("</Attribute>", м.end()) + len("</Attribute>")
    блок = д[початок:кінець]
    новий = re.sub(r'<Attribute uuid="[^"]+">', f'<Attribute uuid="{uuid.uuid4()}">', блок, count=1)
    новий = новий.replace("<Name>СуммаКт</Name>", "<Name>ПодразделениеКазны</Name>")
    новий = новий.replace("<v8:content>Сумма кт</v8:content>", "<v8:content>Подразделение казны</v8:content>")
    новий = новий.replace("<v8:content>Сума Кт</v8:content>", "<v8:content>Підрозділ казни</v8:content>")
    відступ = д[д.rfind("\n", 0, початок) + 1:початок]
    тип = re.search(r"(?P<в>[ \t]*)<Type>[\s\S]*?</Type>", новий)
    в = тип.group("в")
    новий_тип = (f"{в}<Type>{перенос}{в}\t<v8:Type>xs:string</v8:Type>{перенос}"
                 f"{в}\t<v8:StringQualifiers>{перенос}{в}\t\t<v8:Length>150</v8:Length>{перенос}"
                 f"{в}\t\t<v8:AllowedLength>Variable</v8:AllowedLength>{перенос}"
                 f"{в}\t</v8:StringQualifiers>{перенос}{в}</Type>")
    новий = новий[:тип.start()] + новий_тип + новий[тип.end():]
    д = д[:кінець] + перенос + відступ + новий + д[кінець:]
    писати(ДОК, д)
minidom.parse(ДОК)
д = читати(ДОК)
assert д.count("<Name>ПодразделениеКазны</Name>") == 1
м = re.search(r"<Name>ПодразделениеКазны</Name>[\s\S]*?</Type>", д)
assert "xs:string" in м.group(0) and "<v8:Length>150</v8:Length>" in м.group(0)
print("OK: регістр — 10 ресурсів; ТЧ РасшифровкаДокументов + ПодразделениеКазны (Строка 150)")
