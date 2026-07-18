# -*- coding: utf-8 -*-
"""Семантический diff контрольной выгрузки из базы vs наши исходники (BOM/CRLF/хвостовой перевод игнорируем)."""
import sys, difflib
sys.stdout.reconfigure(encoding='utf-8')

TMP = r"C:\Users\SUPPOR~1\AppData\Local\Temp\claude\C--Configuration-downloads-BASERP25--claude-worktrees-1c-composition-document-695999\d31a84b1-2dc5-4bf8-90cf-0d399e911d47\scratchpad\dump_check"
CFG = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"
files = [
    r"Documents\РасчетКомплектаций\Forms\ФормаДокумента\Ext\Form.xml",
    r"Documents\РасчетКомплектаций\Forms\ФормаДокумента\Ext\Form\Module.bsl",
    r"Enums\СтатусыРасчетаКомплектаций.xml",
    r"Documents\РасчетКомплектаций\Templates\МакетАнализСС\Ext\Template.xml",
    r"Documents\РасчетКомплектаций\Templates\МакетПланФакт\Ext\Template.xml",
    r"Documents\РасчетКомплектаций\Templates\МакетПланФактЕтапи\Ext\Template.xml",
    r"Documents\РасчетКомплектаций\Templates\МакетАнализССОдна\Ext\Template.xml",
]


def norm(p):
    return open(p, 'rb').read().decode('utf-8-sig').replace('\r\n', '\n').strip()


for f in files:
    a, b = norm(TMP + "\\" + f), norm(CFG + "\\" + f)
    if a == b:
        print("SEMANT-IDENT:", f)
    else:
        dl = list(difflib.unified_diff(b.splitlines(), a.splitlines(), "наш", "база", lineterm='', n=0))
        print("SEMANT-DIFF :", f, f"({len(dl)} строк диффа)")
        print("\n".join(dl[:14]))
        print("---")
