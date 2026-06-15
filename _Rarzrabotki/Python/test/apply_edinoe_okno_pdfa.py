# -*- coding: utf-8 -*-
# Хирургия форм А_ПриходДенегОтФинАгента до "единого окна" БСП.
# Сохраняет BOM/CRLF каждого файла, проверяет счётчики замен (assert).
import sys, io, os
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

ROOT = r'C:\Configuration_downloads\BASERP25\.claude\worktrees\romantic-burnell-32a095'
DOC_FORM = ROOT + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаДокумента\Ext\Form.xml'
LIST_FORM = ROOT + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаСписка\Ext\Form.xml'
DOC_MODULE = ROOT + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаДокумента\Ext\Form\Module.bsl'
MGR_MODULE = ROOT + r'\Documents\А_ПриходДенегОтФинАгента\Ext\ManagerModule.bsl'
LIST_MODULE = ROOT + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаСписка\Ext\Form\Module.bsl'

def detect(path):
    raw = open(path, 'rb').read()
    bom = raw.startswith(b'\xef\xbb\xbf')
    crlf = b'\r\n' in raw
    return bom, crlf

def read_text(path):
    bom, crlf = detect(path)
    enc = 'utf-8-sig' if bom else 'utf-8'
    with io.open(path, 'r', encoding=enc, newline=None) as f:  # universal newlines -> \n
        return f.read(), bom, crlf

def write_text(path, text, bom, crlf):
    enc = 'utf-8-sig' if bom else 'utf-8'
    nl = '\r\n' if crlf else '\n'
    with io.open(path, 'w', encoding=enc, newline=nl) as f:
        f.write(text)

def must(cond, msg):
    if not cond:
        raise AssertionError(msg)

# ---------- блоки для формы документа ----------
CMDSET = (
"\t<CommandSet>\n"
"\t\t<ExcludedCommand>Post</ExcludedCommand>\n"
"\t\t<ExcludedCommand>PostAndClose</ExcludedCommand>\n"
"\t\t<ExcludedCommand>Write</ExcludedCommand>\n"
"\t</CommandSet>\n"
)

NEW_ACB = (
'\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">\n'
'\t\t<ChildItems>\n'
'\t\t\t<Button name="ФормаПровестиИЗакрыть" id="200">\n'
'\t\t\t\t<Type>CommandBarButton</Type>\n'
'\t\t\t\t<Representation>Text</Representation>\n'
'\t\t\t\t<DefaultButton>true</DefaultButton>\n'
'\t\t\t\t<CommandName>Form.Command.ПровестиИЗакрыть</CommandName>\n'
'\t\t\t\t<ExtendedTooltip name="ФормаПровестиИЗакрытьРасширеннаяПодсказка" id="201"/>\n'
'\t\t\t</Button>\n'
'\t\t\t<Button name="ФормаЗаписать" id="202">\n'
'\t\t\t\t<Type>CommandBarButton</Type>\n'
'\t\t\t\t<Representation>Picture</Representation>\n'
'\t\t\t\t<CommandName>Form.Command.ЗаписатьДокумент</CommandName>\n'
'\t\t\t\t<ExtendedTooltip name="ФормаЗаписатьРасширеннаяПодсказка" id="203"/>\n'
'\t\t\t</Button>\n'
'\t\t\t<Button name="ФормаПровести" id="204">\n'
'\t\t\t\t<Type>CommandBarButton</Type>\n'
'\t\t\t\t<Representation>Picture</Representation>\n'
'\t\t\t\t<CommandName>Form.Command.ПровестиДокумент</CommandName>\n'
'\t\t\t\t<ExtendedTooltip name="ФормаПровестиРасширеннаяПодсказка" id="205"/>\n'
'\t\t\t</Button>\n'
'\t\t\t<Popup name="ПодменюСоздатьНаОсновании" id="206">\n'
'\t\t\t\t<Title>\n'
'\t\t\t\t\t<v8:item>\n'
'\t\t\t\t\t\t<v8:lang>ru</v8:lang>\n'
'\t\t\t\t\t\t<v8:content>Создать на основании</v8:content>\n'
'\t\t\t\t\t</v8:item>\n'
'\t\t\t\t</Title>\n'
'\t\t\t\t<ExtendedTooltip name="ПодменюСоздатьНаОснованииРасширеннаяПодсказка" id="207"/>\n'
'\t\t\t\t<ChildItems>\n'
'\t\t\t\t\t<ButtonGroup name="ПодменюСоздатьНаОснованииВажное" id="208">\n'
'\t\t\t\t\t\t<ExtendedTooltip name="ПодменюСоздатьНаОснованииВажноеРасширеннаяПодсказка" id="209"/>\n'
'\t\t\t\t\t</ButtonGroup>\n'
'\t\t\t\t\t<ButtonGroup name="ПодменюСоздатьНаОснованииОбычное" id="210">\n'
'\t\t\t\t\t\t<ExtendedTooltip name="ПодменюСоздатьНаОснованииОбычноеРасширеннаяПодсказка" id="211"/>\n'
'\t\t\t\t\t</ButtonGroup>\n'
'\t\t\t\t\t<ButtonGroup name="ПодменюСоздатьНаОснованииСмТакже" id="212">\n'
'\t\t\t\t\t\t<ExtendedTooltip name="ПодменюСоздатьНаОснованииСмТакжеРасширеннаяПодсказка" id="213"/>\n'
'\t\t\t\t\t</ButtonGroup>\n'
'\t\t\t\t</ChildItems>\n'
'\t\t\t</Popup>\n'
'\t\t\t<ButtonGroup name="ФормаГлобальныеКоманды" id="214">\n'
'\t\t\t\t<CommandSource>FormCommandPanelGlobalCommands</CommandSource>\n'
'\t\t\t\t<ExtendedTooltip name="ФормаГлобальныеКомандыРасширеннаяПодсказка" id="215"/>\n'
'\t\t\t</ButtonGroup>\n'
'\t\t\t<Popup name="ПодменюПечать" id="216">\n'
'\t\t\t\t<Title>\n'
'\t\t\t\t\t<v8:item>\n'
'\t\t\t\t\t\t<v8:lang>ru</v8:lang>\n'
'\t\t\t\t\t\t<v8:content>Печать</v8:content>\n'
'\t\t\t\t\t</v8:item>\n'
'\t\t\t\t</Title>\n'
'\t\t\t\t<ExtendedTooltip name="ПодменюПечатьРасширеннаяПодсказка" id="217"/>\n'
'\t\t\t</Popup>\n'
'\t\t\t<Popup name="ПодменюОтчеты" id="218">\n'
'\t\t\t\t<Title>\n'
'\t\t\t\t\t<v8:item>\n'
'\t\t\t\t\t\t<v8:lang>ru</v8:lang>\n'
'\t\t\t\t\t\t<v8:content>Отчеты</v8:content>\n'
'\t\t\t\t\t</v8:item>\n'
'\t\t\t\t</Title>\n'
'\t\t\t\t<ExtendedTooltip name="ПодменюОтчетыРасширеннаяПодсказка" id="219"/>\n'
'\t\t\t\t<ChildItems>\n'
'\t\t\t\t\t<ButtonGroup name="ПодменюОтчетыВажное" id="220">\n'
'\t\t\t\t\t\t<ExtendedTooltip name="ПодменюОтчетыВажноеРасширеннаяПодсказка" id="221"/>\n'
'\t\t\t\t\t</ButtonGroup>\n'
'\t\t\t\t\t<ButtonGroup name="ПодменюОтчетыОбычное" id="222">\n'
'\t\t\t\t\t\t<ExtendedTooltip name="ПодменюОтчетыОбычноеРасширеннаяПодсказка" id="223"/>\n'
'\t\t\t\t\t</ButtonGroup>\n'
'\t\t\t\t\t<ButtonGroup name="ПодменюОтчетыСмТакже" id="224">\n'
'\t\t\t\t\t\t<ExtendedTooltip name="ПодменюОтчетыСмТакжеРасширеннаяПодсказка" id="225"/>\n'
'\t\t\t\t\t</ButtonGroup>\n'
'\t\t\t\t</ChildItems>\n'
'\t\t\t</Popup>\n'
'\t\t\t<Popup name="ПодменюЗаполнить" id="226">\n'
'\t\t\t\t<Title>\n'
'\t\t\t\t\t<v8:item>\n'
'\t\t\t\t\t\t<v8:lang>ru</v8:lang>\n'
'\t\t\t\t\t\t<v8:content>Заполнить</v8:content>\n'
'\t\t\t\t\t</v8:item>\n'
'\t\t\t\t</Title>\n'
'\t\t\t\t<ExtendedTooltip name="ПодменюЗаполнитьРасширеннаяПодсказка" id="227"/>\n'
'\t\t\t</Popup>\n'
'\t\t</ChildItems>\n'
'\t</AutoCommandBar>'
)

FORM_EVENTS = (
'\t<Events>\n'
'\t\t<Event name="OnCreateAtServer">ПриСозданииНаСервере</Event>\n'
'\t\t<Event name="OnReadAtServer">ПриЧтенииНаСервере</Event>\n'
'\t\t<Event name="BeforeWriteAtServer">ПередЗаписьюНаСервере</Event>\n'
'\t\t<Event name="FillCheckProcessingAtServer">ОбработкаПроверкиЗаполненияНаСервере</Event>\n'
'\t\t<Event name="OnOpen">ПриОткрытии</Event>\n'
'\t\t<Event name="NotificationProcessing">ОбработкаОповещения</Event>\n'
'\t</Events>'
)

PAGES_EVENTS = (
'\t\t\t<Events>\n'
'\t\t\t\t<Event name="OnCurrentPageChange">ГруппаСтраницыПриСменеСтраницы</Event>\n'
'\t\t\t</Events>\n'
)

NEW_PAGE = (
'\t\t\t\t<Page name="СтраницаДополнительно" id="230">\n'
'\t\t\t\t\t<Title>\n'
'\t\t\t\t\t\t<v8:item>\n'
'\t\t\t\t\t\t\t<v8:lang>ru</v8:lang>\n'
'\t\t\t\t\t\t\t<v8:content>Дополнительно</v8:content>\n'
'\t\t\t\t\t\t</v8:item>\n'
'\t\t\t\t\t</Title>\n'
'\t\t\t\t\t<ExtendedTooltip name="СтраницаДополнительноРасширеннаяПодсказка" id="231"/>\n'
'\t\t\t\t\t<ChildItems>\n'
'\t\t\t\t\t\t<UsualGroup name="ГруппаДополнительныеРеквизиты" id="232">\n'
'\t\t\t\t\t\t\t<Title>\n'
'\t\t\t\t\t\t\t\t<v8:item>\n'
'\t\t\t\t\t\t\t\t\t<v8:lang>ru</v8:lang>\n'
'\t\t\t\t\t\t\t\t\t<v8:content>Дополнительные реквизиты</v8:content>\n'
'\t\t\t\t\t\t\t\t</v8:item>\n'
'\t\t\t\t\t\t\t</Title>\n'
'\t\t\t\t\t\t\t<Group>Vertical</Group>\n'
'\t\t\t\t\t\t\t<Behavior>Usual</Behavior>\n'
'\t\t\t\t\t\t\t<Representation>NormalSeparation</Representation>\n'
'\t\t\t\t\t\t\t<ShowTitle>false</ShowTitle>\n'
'\t\t\t\t\t\t\t<ExtendedTooltip name="ГруппаДополнительныеРеквизитыРасширеннаяПодсказка" id="233"/>\n'
'\t\t\t\t\t\t</UsualGroup>\n'
'\t\t\t\t\t</ChildItems>\n'
'\t\t\t\t</Page>\n'
)

COMMANDS = (
'\t<Commands>\n'
'\t\t<Command name="ПровестиИЗакрыть" id="1">\n'
'\t\t\t<Title>\n'
'\t\t\t\t<v8:item>\n'
'\t\t\t\t\t<v8:lang>ru</v8:lang>\n'
'\t\t\t\t\t<v8:content>Провести и закрыть</v8:content>\n'
'\t\t\t\t</v8:item>\n'
'\t\t\t</Title>\n'
'\t\t\t<Shortcut>Ctrl+Enter</Shortcut>\n'
'\t\t\t<Action>ПровестиИЗакрыть</Action>\n'
'\t\t\t<ModifiesSavedData>true</ModifiesSavedData>\n'
'\t\t\t<CurrentRowUse>DontUse</CurrentRowUse>\n'
'\t\t</Command>\n'
'\t\t<Command name="ЗаписатьДокумент" id="2">\n'
'\t\t\t<Title>\n'
'\t\t\t\t<v8:item>\n'
'\t\t\t\t\t<v8:lang>ru</v8:lang>\n'
'\t\t\t\t\t<v8:content>Записать</v8:content>\n'
'\t\t\t\t</v8:item>\n'
'\t\t\t</Title>\n'
'\t\t\t<Shortcut>Ctrl+S</Shortcut>\n'
'\t\t\t<Picture>\n'
'\t\t\t\t<xr:Ref>StdPicture.Write</xr:Ref>\n'
'\t\t\t\t<xr:LoadTransparent>true</xr:LoadTransparent>\n'
'\t\t\t</Picture>\n'
'\t\t\t<Action>ЗаписатьДокумент</Action>\n'
'\t\t\t<ModifiesSavedData>true</ModifiesSavedData>\n'
'\t\t\t<CurrentRowUse>DontUse</CurrentRowUse>\n'
'\t\t</Command>\n'
'\t\t<Command name="ПровестиДокумент" id="3">\n'
'\t\t\t<Title>\n'
'\t\t\t\t<v8:item>\n'
'\t\t\t\t\t<v8:lang>ru</v8:lang>\n'
'\t\t\t\t\t<v8:content>Провести</v8:content>\n'
'\t\t\t\t</v8:item>\n'
'\t\t\t</Title>\n'
'\t\t\t<Picture>\n'
'\t\t\t\t<xr:Ref>StdPicture.Post</xr:Ref>\n'
'\t\t\t\t<xr:LoadTransparent>true</xr:LoadTransparent>\n'
'\t\t\t</Picture>\n'
'\t\t\t<Action>ПровестиДокумент</Action>\n'
'\t\t\t<ModifiesSavedData>true</ModifiesSavedData>\n'
'\t\t\t<CurrentRowUse>DontUse</CurrentRowUse>\n'
'\t\t</Command>\n'
'\t</Commands>\n'
)

# ---------- ФОРМА ДОКУМЕНТА ----------
text, bom, crlf = read_text(DOC_FORM)
print('DOC_FORM bom=%s crlf=%s len=%d' % (bom, crlf, len(text)))

# 1. CommandSet перед <AutoTime>
a = '\t<AutoTime>CurrentOrLast</AutoTime>'
must(text.count(a) == 1, '1: AutoTime anchor count != 1')
text = text.replace(a, CMDSET + a, 1)

# 2. Заменить form-level AutoCommandBar + добавить Events (index-based)
acb_open = '<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">'
i = text.index(acb_open)
ls = text.rindex('\n', 0, i) + 1
j = text.index('</AutoCommandBar>', i) + len('</AutoCommandBar>')
text = text[:ls] + NEW_ACB + '\n' + FORM_EVENTS + text[j:]
print('2: AutoCommandBar replaced')

# 3. Удалить меню-контейнеры из тела (контигуальный блок 159..177 -> до ГруппаНомерДата)
s = text.index('<UsualGroup name="ПодменюСоздатьНаОсновании" id="159">')
ls2 = text.rindex('\n', 0, s) + 1
e = text.index('<UsualGroup name="ГруппаНомерДата" id="1">')
le2 = text.rindex('\n', 0, e) + 1
removed = text[ls2:le2]
must('id="167"' in removed and 'id="177"' in removed and 'id="1"' not in removed.split('ГруппаНомерДата')[0][-40:], '3: body-menu span suspicious')
text = text[:ls2] + text[le2:]
print('3: body menu groups removed (%d chars)' % len(removed))

# 4. Events группы страниц (OnCurrentPageChange)
pg = '\t\t\t<ExtendedTooltip name="ГруппаСтраницыРасширеннаяПодсказка" id="15"/>\n\t\t\t<ChildItems>'
must(text.count(pg) == 1, '4: Pages anchor count != 1')
text = text.replace(pg, '\t\t\t<ExtendedTooltip name="ГруппаСтраницыРасширеннаяПодсказка" id="15"/>\n' + PAGES_EVENTS + '\t\t\t<ChildItems>', 1)

# 5. Новая страница "Дополнительно" перед закрытием группы страниц
close_pages = '\t\t\t</ChildItems>\n\t\t</Pages>'
must(text.count(close_pages) == 1, '5: Pages close anchor count != 1')
text = text.replace(close_pages, NEW_PAGE + close_pages, 1)

# 6. Commands перед </Form>
endf = '\t</Attributes>\n</Form>'
must(text.count(endf) == 1, '6: Attributes/Form anchor count != 1')
text = text.replace(endf, '\t</Attributes>\n' + COMMANDS + '</Form>', 1)

write_text(DOC_FORM, text, bom, crlf)
print('DOC_FORM written, new len=%d' % len(text))

# ---------- ФОРМА СПИСКА: вставить CommandSet ----------
LIST_CMDSET = (
"\t<CommandSet>\n" +
"".join("\t\t<ExcludedCommand>%s</ExcludedCommand>\n" % c for c in [
 "Abort","Cancel","CancelSearch","Change","Copy","Create","Delete","DynamicListStandardSettings",
 "Find","FindByCurrentValue","Ignore","ListSettings","LoadDynamicListSettings","No","OK","OutputList",
 "Post","Refresh","RestoreValues","Retry","SaveDynamicListSettings","SaveValues","SetDateInterval",
 "SetDeletionMark","UndoPosting","Yes"]) +
"\t</CommandSet>\n"
)
ltext, lbom, lcrlf = read_text(LIST_FORM)
print('LIST_FORM bom=%s crlf=%s len=%d' % (lbom, lcrlf, len(ltext)))
la = '\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">'
must(ltext.count(la) == 1, 'L: list AutoCommandBar anchor count != 1')
must('<CommandSet>' not in ltext, 'L: list already has CommandSet')
ltext = ltext.replace(la, LIST_CMDSET + la, 1)
write_text(LIST_FORM, ltext, lbom, lcrlf)
print('LIST_FORM written, new len=%d' % len(ltext))

# ---------- ФИКСАП КОДИРОВКИ НОВЫХ .bsl (под конвенцию существующего модуля списка) ----------
ref_bom, ref_crlf = detect(LIST_MODULE)
print('Reference .bsl convention: bom=%s crlf=%s' % (ref_bom, ref_crlf))
for p in (DOC_MODULE, MGR_MODULE):
    t, b, c = read_text(p)
    has_tab = '\t' in t
    write_text(p, t, ref_bom, ref_crlf)
    print('FIXED %s : had_tab=%s -> bom=%s crlf=%s' % (os.path.basename(os.path.dirname(os.path.dirname(p))) + '/' + os.path.basename(p), has_tab, ref_bom, ref_crlf))

print('OK ALL DONE')
