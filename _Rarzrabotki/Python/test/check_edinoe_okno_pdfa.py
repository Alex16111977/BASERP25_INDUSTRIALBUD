# -*- coding: utf-8 -*-
import sys, re, xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')
R = r'C:\Configuration_downloads\BASERP25\.claude\worktrees\romantic-burnell-32a095'
docf = R + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаДокумента\Ext\Form.xml'
listf = R + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаСписка\Ext\Form.xml'
mod = R + r'\Documents\А_ПриходДенегОтФинАгента\Forms\ФормаДокумента\Ext\Form\Module.bsl'

for f in (docf, listf):
    ET.parse(f)
    print('WELLFORMED', f.split('Forms')[1])

xml = open(docf, encoding='utf-8-sig').read()
modtext = open(mod, encoding='utf-8-sig').read()
events = re.findall(r'<Event name="[^"]+">([^<]+)</Event>', xml)
actions = re.findall(r'<Action>([^<]+)</Action>', xml)
procs = set(re.findall(r'Процедура\s+([0-9A-Za-zА-Яа-я_]+)\s*\(', modtext))
print('EVENTS', events)
print('ACTIONS', actions)
print('MISSING_PROC', [e for e in events + actions if e not in procs])

def span(s, a, b):
    i = s.find(a); j = s.find(b)
    return s[i:j] if i >= 0 and j >= 0 else ''

item_xml = xml.replace(span(xml, '<Attributes>', '</Attributes>'), '').replace(span(xml, '<Commands>', '</Commands>'), '')
ids = re.findall(r' id="(-?\d+)"', item_xml)
dups = sorted({x for x in ids if ids.count(x) > 1})
print('ITEM_ID_DUPLICATES', dups)
print('NEW_IDS_PRESENT', all(('id="%d"' % n) in xml for n in [200, 206, 214, 218, 226, 230, 232]))
print('BODY_MENU_GONE', '"ПодменюОтчеты" id="169"' not in xml and 'id="159"' not in xml)
# контейнеры меню теперь ДО </AutoCommandBar>
acb_close = xml.index('</AutoCommandBar>')
print('MENU_IN_ACB', xml.index('name="ПодменюОтчеты"') < acb_close)
print('LIST_HAS_COMMANDSET', '<CommandSet>' in open(listf, encoding='utf-8-sig').read())
print('DONE')
