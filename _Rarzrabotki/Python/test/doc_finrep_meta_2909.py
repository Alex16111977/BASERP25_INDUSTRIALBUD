# -*- coding: utf-8 -*-
"""Метадані частини 2 (29.09.2026): регістр +3 ресурси (строк оплати, база строку, не показувати),
ТЧ РасшифровкаДокументов + ДатаПодписания, ТЧ Расчет + ОстатокАванса, форма — команда «Зберегти в Excel».
Ідемпотентний. Після запису — minidom і assert."""
import io, re, sys, uuid
from xml.dom import minidom
sys.stdout.reconfigure(encoding="utf-8")
# шаблони ресурсу — копія з doc_finrep_meta_2809.py (той скрипт при імпорті виконується)
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



ФОРМА = ДОК.replace(".xml", r"\Forms\ФормаДокумента\Ext\Form.xml")
ЧИСЛО3 = """<v8:Type>xs:decimal</v8:Type>
					<v8:NumberQualifiers>
						<v8:Digits>3</v8:Digits>
						<v8:FractionDigits>0</v8:FractionDigits>
						<v8:AllowedSign>Nonnegative</v8:AllowedSign>
					</v8:NumberQualifiers>"""
РЯДОК20 = """<v8:Type>xs:string</v8:Type>
					<v8:StringQualifiers>
						<v8:Length>20</v8:Length>
						<v8:AllowedLength>Variable</v8:AllowedLength>
					</v8:StringQualifiers>"""
БУЛЕВО = "<v8:Type>xs:boolean</v8:Type>"
РЕСУРСИ = [
    ("СрокОплатыДней", "Срок оплаты, дней", "Строк оплати, днів", ЧИСЛО3),
    ("БазаСрокаОплаты", "База срока оплаты", "Від чого рахувати строк оплати", РЯДОК20),
    ("НеПоказыватьВОтчете", "Не показывать в отчете", "Не показувати у звіті", БУЛЕВО),
]

# --- регістр
т = читати(РЕГ)
перенос = "\r\n" if "\r\n" in т else "\n"
нові = "".join(ресурс(*р) for р in РЕСУРСИ if f"<Name>{р[0]}</Name>" not in т)
if нові:
    якір = "\t\t\t<Dimension uuid="
    assert т.count(якір) == 1, "якір Dimension"
    т = т.replace(якір, нові.replace("\n", перенос) + якір, 1)
    писати(РЕГ, т)
minidom.parse(РЕГ)
т = читати(РЕГ)
assert т.count("<Resource uuid=") == 13, т.count("<Resource uuid=")


def клон(д, зразок, імя, ru, uk, тип_xml=None):
    """Копія реквізиту ТЧ з іменем-зразком (унікальним у файлі) під новим ім'ям одразу після нього."""
    if f"<Name>{імя}</Name>" in д:
        return д
    assert д.count(f"<Name>{зразок}</Name>") == 1, зразок
    м = re.search(f"<Name>{зразок}</Name>", д)
    початок = д.rfind("<Attribute uuid=", 0, м.start())
    кінець = д.find("</Attribute>", м.end()) + len("</Attribute>")
    новий = re.sub(r'<Attribute uuid="[^"]+">', f'<Attribute uuid="{uuid.uuid4()}">', д[початок:кінець], count=1)
    новий = новий.replace(f"<Name>{зразок}</Name>", f"<Name>{імя}</Name>")
    новий = re.sub(r"(<v8:lang>ru</v8:lang>\s*<v8:content>)[^<]*", lambda x: x.group(1) + ru, новий, count=1)
    новий = re.sub(r"(<v8:lang>uk</v8:lang>\s*<v8:content>)[^<]*", lambda x: x.group(1) + uk, новий, count=1)
    if тип_xml:
        тип = re.search(r"(?P<в>[ \t]*)<Type>[\s\S]*?</Type>", новий)
        в = тип.group("в")
        новий = новий[:тип.start()] + f"{в}<Type>{перенос}{в}\t{тип_xml}{перенос}{в}</Type>" + новий[тип.end():]
    відступ = д[д.rfind("\n", 0, початок) + 1:початок]
    return д[:кінець] + перенос + відступ + новий + д[кінець:]


д = читати(ДОК)
д = клон(д, "ДатаДокумента", "ДатаПодписания", "Дата подписания", "Дата підписання")
д = клон(д, "ИтогоКПолучению", "ОстатокАванса", "Остаток аванса", "Залишок авансу")
писати(ДОК, д)
minidom.parse(ДОК)
д = читати(ДОК)
assert д.count("<Name>ДатаПодписания</Name>") == 1 and д.count("<Name>ОстатокАванса</Name>") == 1

# --- форма: команда і кнопка «Зберегти в Excel» після «Зберегти динаміку в HTML»
ф = читати(ФОРМА)
if 'name="СохранитьExcel"' not in ф:
    ід = [int(x) for x in re.findall(r'<Command name="[^"]+" id="(\d+)"', ф)]
    ід_ком = max(ід) + 1
    ід_ел = max(int(x) for x in re.findall(r'<(?!Command)\w+ name="[^"]+" id="(\d+)"', ф)) + 1
    заголовок = ("<Title>\n\t\t\t\t<v8:item>\n\t\t\t\t\t<v8:lang>ru</v8:lang>\n\t\t\t\t\t<v8:content>Зберегти в Excel"
                 "</v8:content>\n\t\t\t\t</v8:item>\n\t\t\t\t<v8:item>\n\t\t\t\t\t<v8:lang>uk</v8:lang>\n\t\t\t\t\t"
                 "<v8:content>Зберегти в Excel</v8:content>\n\t\t\t\t</v8:item>\n\t\t\t</Title>")
    команда = (f'\t\t<Command name="СохранитьExcel" id="{ід_ком}">\n\t\t\t{заголовок}\n'
               f'\t\t\t<Action>СохранитьExcel</Action>\n\t\t</Command>\n')
    якір_к = '\t\t<Command name="СбросФильтраРасшифровки"'
    assert ф.count(якір_к) == 1
    ф = ф.replace(якір_к, команда.replace("\n", перенос) + якір_к, 1)
    кнопка_старт = ф.index('<Button name="КнопкаСохранитьДинамику"')
    кнопка_кінець = ф.index("</Button>", кнопка_старт) + len("</Button>")
    відступ = ф[ф.rfind("\n", 0, кнопка_старт) + 1:кнопка_старт]
    т6 = "\t" * 6
    кнопка = (f'<Button name="КнопкаСохранитьExcel" id="{ід_ел}">\n{т6}\t<Type>UsualButton</Type>\n'
              f'{т6}\t<CommandName>Form.Command.СохранитьExcel</CommandName>\n'
              f'{т6}\t<ExtendedTooltip name="КнопкаСохранитьExcelРасширеннаяПодсказка" id="{ід_ел + 1}"/>\n'
              f'{т6}</Button>')
    ф = ф[:кнопка_кінець] + перенос + відступ + кнопка.replace("\n", перенос) + ф[кнопка_кінець:]
    писати(ФОРМА, ф)
minidom.parse(ФОРМА)
assert читати(ФОРМА).count('name="КнопкаСохранитьExcel"') == 1
print("OK: регістр 13 ресурсів; ТЧ +ДатаПодписания, +ОстатокАванса; форма + «Зберегти в Excel»")
