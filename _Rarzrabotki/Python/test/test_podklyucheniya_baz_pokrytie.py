# -*- coding: utf-8 -*-
"""Покрытие: каждый идентификатор базы, встречающийся в коде (конфигурация + _Rarzrabotki),
есть в регистре А_НастройкиПодключенияБаз и даёт живое COM-соединение.
"""
import os, re, sys
import win32com.client
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONFIG_ROOT = r"C:\Configuration_downloads\BASERP25"
RE_CALL = re.compile(r'А_ПодключенияБазСервер\.(?:УстановитьСоединение|СтрокаПодключения)\("([^"]+)"\)')

def err_text(e):
    return (e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)).strip()[:160]

# 1. Собрать все идентификаторы из кода
used = {}
SKIP = os.path.join("_Rarzrabotki", "Рабочие места", "GitHub")  # чужие репозитории, не наш код

for sub in ["Documents", "Reports", "CommonModules", "_Rarzrabotki"]:
    for dp, _, files in os.walk(os.path.join(CONFIG_ROOT, sub)):
        if SKIP in dp:
            continue
        for f in files:
            if not f.endswith(".bsl"):
                continue
            p = os.path.join(dp, f)
            try:
                txt = open(p, encoding="utf-8").read()
            except Exception:
                continue
            for m in RE_CALL.finditer(txt):
                used.setdefault(m.group(1), 0)
                used[m.group(1)] += 1

print("Идентификаторы, используемые в коде:")
for k in sorted(used):
    print(f"  {k}: {used[k]} вызовов")

# 2. Хардкода быть не должно
hard = 0
for sub in ["Documents", "Reports", "_Rarzrabotki"]:
    for dp, _, files in os.walk(os.path.join(CONFIG_ROOT, sub)):
        if SKIP in dp:
            continue
        for f in files:
            if f.endswith(".bsl"):
                try:
                    if "SQLSERVER" in open(os.path.join(dp, f), encoding="utf-8").read():
                        print(f"  !!! ХАРДКОД SQLSERVER: {os.path.join(dp, f)}")
                        hard += 1
                except Exception:
                    pass
print(f"\nФайлов с хардкодом SQLSERVER: {hard} (ожидаем 0)")

# 3. Каждый идентификатор → живое соединение
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
M = erp.А_ПодключенияБазСервер

fails = hard
for base in sorted(used):
    try:
        conn = M.УстановитьСоединение(base)
        print(f"OK {base}: {conn.Metadata.Synonym}")
        conn = None
    except Exception as e:
        print(f"FAIL {base}: {err_text(e)}")
        fails += 1

print("ИТОГ:", "OK" if fails == 0 else f"FAIL ({fails})")
