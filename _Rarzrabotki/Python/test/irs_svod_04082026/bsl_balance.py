# -*- coding: utf-8 -*-
"""Перевірка балансу блоків BSL: Якщо/Цикл/Спроба/Процедура/Функція/Область."""
import io, re, sys
sys.stdout.reconfigure(encoding='utf-8')

PAIRS = [
    (r"^\s*Если\b(?!.*\bТогда\s*(//.*)?$)", None),  # placeholder, replaced below
]

def check(path):
    s = io.open(path, encoding="utf-8").read()
    lines = s.split("\n")
    def cnt(pat):
        return sum(1 for ln in lines if re.match(pat, ln, re.I))
    checks = [
        ("Если",       r"\s*Если\b",         "КонецЕсли",     r"\s*КонецЕсли\s*;"),
        ("Для/Пока",   r"\s*(Для|Пока)\b",   "КонецЦикла",    r"\s*КонецЦикла\s*;"),
        ("Попытка",    r"\s*Попытка\b",      "КонецПопытки",  r"\s*КонецПопытки\s*;"),
        ("Процедура",  r"\s*Процедура\b",    "КонецПроцедуры",r"\s*КонецПроцедуры\b"),
        ("Функция",    r"\s*Функция\b",      "КонецФункции",  r"\s*КонецФункции\b"),
        ("#Область",   r"#Область\b",        "#КонецОбласти", r"#КонецОбласти\b"),
        ("#Если",      r"#Если\b",           "#КонецЕсли",    r"#КонецЕсли\b"),
    ]
    print("=== %s ===" % path)
    ok = True
    for na, pa, nb, pb in checks:
        a, b = cnt(pa), cnt(pb)
        good = (a == b)
        ok = ok and good
        print("  %-12s %3d  vs  %-16s %3d   %s" % (na, a, nb, b, "OK" if good else "!!! РОЗБІЖНІСТЬ"))
    # типові одруківки
    bad = []
    for i, ln in enumerate(lines, 1):
        for w in ("КонецФункция", "КонецПроцедура", "КонецЕслі", "КонецЦікла", "КонецПопитки"):
            if w in ln:
                bad.append((i, w, ln.strip()[:60]))
    print("  одруківки: %s" % (bad if bad else "немає"))
    # заборонені імена
    res = [(i, ln.strip()[:60]) for i, ln in enumerate(lines, 1)
           if re.match(r"\s*(Перем\s+)?Ссылка\s*=", ln)]
    print("  змінна 'Ссылка': %s" % (res if res else "немає"))
    return ok and not bad


import glob
allok = True
for p in sys.argv[1:]:
    for g in glob.glob(p):
        allok = check(g) and allok
print("\n%s" % ("BSL БАЛАНС OK" if allok else "!!! Є ПРОБЛЕМИ !!!"))
