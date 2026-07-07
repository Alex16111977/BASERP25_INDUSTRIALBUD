# -*- coding: utf-8 -*-
# Упаковка P.epf против самодостаточного стаб-IB (BuhBud-типы СписаниеТоваров/
# ПоступлениеТоваровУслуг/РеализацияТоваровУслуг). Реальную BuhBud не трогаем —
# .epf потом компилируется против настоящих типов при ВнешниеОбработки.Создать в Enterprise.
import os, re, shutil, subprocess, uuid, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

PLAT = r'C:\Program Files\1cv8\8.3.20.1914\bin\1cv8.exe'
STUB = r'C:\Users\support1c\AppData\Local\Temp\epf_stub_db_1682166821\cfg'
WORK = r'C:\1ctmp\stubwork'
IB   = r'C:\1ctmp\stubib'
LOG  = r'C:\1ctmp\stub_build.log'
SRC  = r'C:\1ctmp\zpr\P.xml'
OUT  = r'C:\1ctmp\zpr\P.epf'
DEST = r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Заполнить в бухучете приобритение по списанию.epf'
NEW_DOCS = ['СписаниеТоваров', 'ПоступлениеТоваровУслуг', 'РеализацияТоваровУслуг', 'ПеремещениеТоваров']
GUID_RE = re.compile(r'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}')

def run1c(args, tag):
    if os.path.exists(LOG):
        os.remove(LOG)
    p = subprocess.run([PLAT] + args, capture_output=True)
    log = ''
    if os.path.exists(LOG):
        log = open(LOG, encoding='utf-8-sig', errors='replace').read()
    extra = ''
    if not log.strip():
        so = (p.stdout or b'').decode('utf-8', 'replace').strip()
        se = (p.stderr or b'').decode('utf-8', 'replace').strip()
        extra = (so + ' ' + se).strip()
    print(f"[{tag}] exit={p.returncode}  {log.strip()[:400]} {extra[:300]}")
    return p.returncode

# 1. подготовка рабочей копии стаба
if os.path.exists(WORK): shutil.rmtree(WORK)
if os.path.exists(IB):   shutil.rmtree(IB)
shutil.copytree(STUB, WORK)
os.makedirs(r'C:\1ctmp', exist_ok=True)  # IB-каталог создаёт сама платформа

# 2. патч версий 2.17->2.13, 8_3_24->8_3_20
patched = 0
for root, _, files in os.walk(WORK):
    for f in files:
        if f.endswith('.xml'):
            fp = os.path.join(root, f)
            t = open(fp, encoding='utf-8-sig').read()
            t2 = t.replace('version="2.17"', 'version="2.13"').replace('Version8_3_24', 'Version8_3_20')
            if t2 != t:
                open(fp, 'w', encoding='utf-8-sig').write(t2); patched += 1
print(f"patched xml: {patched}")

# 3. читаем шаблон, затем чистим стаб до минимума (нужны только 3 документа)
tmpl = open(os.path.join(WORK, 'Documents', 'КорректировкаРегистров.xml'), encoding='utf-8-sig').read()
for d in ['Catalogs', 'ChartsOfAccounts', 'ChartsOfCharacteristicTypes']:
    shutil.rmtree(os.path.join(WORK, d), ignore_errors=True)
for f in ['ОперацияБух.xml', 'КорректировкаРегистров.xml']:
    fp = os.path.join(WORK, 'Documents', f)
    if os.path.exists(fp): os.remove(fp)

# 4. клон шаблона в 3 документа (regen всех GUID + переименование)
for name in NEW_DOCS:
    body = tmpl.replace('КорректировкаРегистров', name)
    body = GUID_RE.sub(lambda m: str(uuid.uuid4()), body)
    open(os.path.join(WORK, 'Documents', name + '.xml'), 'w', encoding='utf-8-sig').write(body)
print(f"cloned docs: {NEW_DOCS}")

# 5. минимальный ChildObjects в Configuration.xml: Язык + 3 документа
cfg = os.path.join(WORK, 'Configuration.xml')
c = open(cfg, encoding='utf-8-sig').read()
docs = ''.join(f'\t\t\t<Document>{n}</Document>\n' for n in NEW_DOCS)
new_block = '<ChildObjects>\n\t\t\t<Language>Русский</Language>\n' + docs + '\t\t</ChildObjects>'
c2 = re.sub(r'<ChildObjects>.*?</ChildObjects>', new_block, c, flags=re.S)
assert c2 != c, "ChildObjects не заменён"
open(cfg, 'w', encoding='utf-8-sig').write(c2)
print("Configuration.xml: минимизирован (Язык + 3 документа)")

# 5. создать IB, загрузить конфигурацию, обновить, упаковать
cs = f'File={IB};'
if run1c(['CREATEINFOBASE', cs, '/DisableStartupDialogs'], 'CREATE') != 0: sys.exit(1)
if run1c(['DESIGNER', '/IBConnectionString', cs, '/LoadConfigFromFiles', WORK, '/Out', LOG, '/DisableStartupDialogs'], 'LOADCFG') != 0: sys.exit(1)
run1c(['DESIGNER', '/IBConnectionString', cs, '/UpdateDBCfg', '/Out', LOG, '/DisableStartupDialogs'], 'UPDATEDB')
if os.path.exists(OUT): os.remove(OUT)
run1c(['DESIGNER', '/IBConnectionString', cs, '/LoadExternalDataProcessorOrReportFromFiles', SRC, OUT, '/Out', LOG, '/DisableStartupDialogs'], 'PACK')

# 6. результат
if os.path.exists(OUT) and os.path.getsize(OUT) > 0:
    shutil.copyfile(OUT, DEST)
    print(f"OK EPF built: {os.path.getsize(OUT)} bytes -> {DEST}")
else:
    print("FAIL: EPF не создан")
    sys.exit(1)
