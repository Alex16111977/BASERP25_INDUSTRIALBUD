# -*- coding: utf-8 -*-
# Откат «Свойств» и страницы «Дополнительно» в форме документа А_ПриходДенегОтФинАгента.
import sys, io, os
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

ROOT = r'C:\Configuration_downloads\BASERP25\.claude\worktrees\romantic-burnell-32a095'
DOC_FORM = ROOT + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаДокумента\Ext\Form.xml'
DOC_MODULE = ROOT + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаДокумента\Ext\Form\Module.bsl'
MGR_MODULE = ROOT + r'\Documents\А_ПриходДенегОтФинАгента\Ext\ManagerModule.bsl'
LIST_MODULE = ROOT + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаСписка\Ext\Form\Module.bsl'

def detect(path):
    raw = open(path, 'rb').read()
    return raw.startswith(b'\xef\xbb\xbf'), (b'\r\n' in raw)

def read_text(path):
    bom, crlf = detect(path)
    return io.open(path, 'r', encoding='utf-8-sig' if bom else 'utf-8', newline=None).read(), bom, crlf

def write_text(path, text, bom, crlf):
    io.open(path, 'w', encoding='utf-8-sig' if bom else 'utf-8', newline='\r\n' if crlf else '\n').write(text)

def must(c, m):
    if not c: raise AssertionError(m)

text, bom, crlf = read_text(DOC_FORM)
print('DOC_FORM bom=%s crlf=%s len=%d' % (bom, crlf, len(text)))

# 1) Свернуть form-level <Events> до одного OnCreateAtServer
OLD_EVENTS = (
'\t<Events>\n'
'\t\t<Event name="OnCreateAtServer">ПриСозданииНаСервере</Event>\n'
'\t\t<Event name="OnReadAtServer">ПриЧтенииНаСервере</Event>\n'
'\t\t<Event name="BeforeWriteAtServer">ПередЗаписьюНаСервере</Event>\n'
'\t\t<Event name="FillCheckProcessingAtServer">ОбработкаПроверкиЗаполненияНаСервере</Event>\n'
'\t\t<Event name="OnOpen">ПриОткрытии</Event>\n'
'\t\t<Event name="NotificationProcessing">ОбработкаОповещения</Event>\n'
'\t</Events>'
)
NEW_EVENTS = (
'\t<Events>\n'
'\t\t<Event name="OnCreateAtServer">ПриСозданииНаСервере</Event>\n'
'\t</Events>'
)
must(text.count(OLD_EVENTS) == 1, '1: form Events block not found uniquely')
text = text.replace(OLD_EVENTS, NEW_EVENTS, 1)
print('1: form-level Events collapsed to OnCreateAtServer')

# 2) Удалить Events группы страниц (OnCurrentPageChange)
PAGES_EVENTS = (
'\t\t\t<Events>\n'
'\t\t\t\t<Event name="OnCurrentPageChange">ГруппаСтраницыПриСменеСтраницы</Event>\n'
'\t\t\t</Events>\n'
)
must(text.count(PAGES_EVENTS) == 1, '2: pages Events block not found uniquely')
text = text.replace(PAGES_EVENTS, '', 1)
print('2: pages OnCurrentPageChange removed')

# 3) Удалить страницу "Дополнительно" целиком
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
must(text.count(NEW_PAGE) == 1, '3: Дополнительно page block not found uniquely')
text = text.replace(NEW_PAGE, '', 1)
print('3: СтраницаДополнительно removed')

# guard: остались только нужные события и нет следов Свойств
must('СтраницаДополнительно' not in text, 'guard: Дополнительно still present')
must('ГруппаДополнительныеРеквизиты' not in text, 'guard: ГруппаДополнительныеРеквизиты still present')
must('OnCurrentPageChange' not in text, 'guard: OnCurrentPageChange still present')
must('OnReadAtServer' not in text and 'NotificationProcessing' not in text, 'guard: stray events remain')
# командная панель и команды формы — на месте
must('<CommandSet>' in text and 'Form.Command.ПровестиИЗакрыть' in text and '<Commands>' in text, 'guard: core panel/commands lost')
write_text(DOC_FORM, text, bom, crlf)
print('DOC_FORM written, new len=%d' % len(text))

# фиксап кодировки переписанных .bsl
rb, rc = detect(LIST_MODULE)
for p in (DOC_MODULE, MGR_MODULE):
    t, b, c = read_text(p)
    write_text(p, t, rb, rc)
    print('FIXED %s had_tab=%s' % (os.path.basename(p), '\t' in t))

print('OK DONE')
