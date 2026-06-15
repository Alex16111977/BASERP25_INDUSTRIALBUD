# -*- coding: utf-8 -*-
import sys, re, xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')
R = r'C:\Configuration_downloads\BASERP25\.claude\worktrees\romantic-burnell-32a095'
df = R + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаДокумента\Ext\Form.xml'
mf = R + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаДокумента\Ext\Form\Module.bsl'
mg = R + r'\Documents\А_ПриходДенегОтФинАгента\Ext\ManagerModule.bsl'
ET.parse(df)
x = open(df, encoding='utf-8-sig').read()
m = open(mf, encoding='utf-8-sig').read()
g = open(mg, encoding='utf-8-sig').read()
ev = re.findall(r'<Event name="[^"]+">([^<]+)</Event>', x)
ac = re.findall(r'<Action>([^<]+)</Action>', x)
pr = set(re.findall(r'Процедура\s+([0-9A-Za-zА-Яа-я_]+)\s*\(', m))
print('WELLFORMED ok')
print('EVENTS', ev)
print('ACTIONS', ac)
print('MISSING_PROC', [e for e in ev + ac if e not in pr])
print('NO_SVOYSTVA_IN_MODULE', ('УправлениеСвойствами' not in m) and ('ПараметрыСвойств' not in m))
print('NO_SVOYSTVA_IN_FORM', ('СтраницаДополнительно' not in x) and ('ГруппаДополнительныеРеквизиты' not in x))
print('CORE_PANEL_OK', ('<CommandSet>' in x) and ('Form.Command.ПровестиИЗакрыть' in x) and ('name="ПодменюОтчеты"' in x))
print('MGR_NO_OTCHETY', 'ДобавитьКомандыОтчетов' not in g, '| MGR_HAS_VERSION', 'ПриОпределенииНастроекВерсионированияОбъектов' in g)
