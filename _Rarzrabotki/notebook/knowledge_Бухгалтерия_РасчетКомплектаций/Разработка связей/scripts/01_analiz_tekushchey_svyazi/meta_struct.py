import sys, re
import xml.etree.ElementTree as ET
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
NS = {'md':'http://v8.1c.ru/8.3/MDClasses','v8':'http://v8.1c.ru/8.1/data/core','xr':'http://v8.1c.ru/8.3/xcf/readable'}
def syn(el):
    s = el.find('md:Properties/md:Synonym', NS)
    if s is None: return ''
    out=[]
    for it in s.findall('v8:item', NS):
        out.append(it.find('v8:lang',NS).text+':'+(it.find('v8:content',NS).text or ''))
    return '; '.join(out)
def typ(el):
    t = el.find('md:Properties/md:Type', NS)
    if t is None: return ''
    parts=[]
    for x in t:
        tag = x.tag.split('}')[1]
        if tag=='Type': parts.append(x.text)
        elif tag=='StringQualifiers':
            parts.append('len=' + (x.find('v8:Length',NS).text))
        elif tag=='NumberQualifiers':
            parts.append('num=' + x.find('v8:Digits',NS).text + ',' + x.find('v8:FractionDigits',NS).text)
    return ' '.join(parts)
for f in sys.argv[1:]:
    root = ET.parse(f).getroot()
    obj = root[0]
    print('=== ', obj.find('md:Properties/md:Name',NS).text, '|', syn(obj))
    co = obj.find('md:ChildObjects', NS)
    if co is None: continue
    for ch in co:
        tag = ch.tag.split('}')[1]
        if tag in ('Attribute','Dimension','Resource'):
            print(f"  {tag}: {ch.find('md:Properties/md:Name',NS).text} [{typ(ch)}] {syn(ch)}")
        elif tag=='TabularSection':
            print(f"  TS: {ch.find('md:Properties/md:Name',NS).text} {syn(ch)}")
            for a in ch.find('md:ChildObjects',NS):
                print(f"     - {a.find('md:Properties/md:Name',NS).text} [{typ(a)}] {syn(a)}")
