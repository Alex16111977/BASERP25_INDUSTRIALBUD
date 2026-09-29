# -*- coding: utf-8 -*-
"""Метадані 29.09.2026 (б): регістр +УсловияДоговора Строка(1000), +СтатусВОтчете Строка(10)
(«» авто / «Всегда» / «Закрыт»); ресурс НеПоказыватьВОтчете прибрано (його роль — статус «Закрыт»).
Ідемпотентний."""
import io, re, sys, uuid
from xml.dom import minidom
sys.stdout.reconfigure(encoding="utf-8")
src = io.open(__file__.replace("doc_finrep_meta_2909b.py", "doc_finrep_meta_2809.py"), encoding="utf-8").read()
exec(src[src.index("КОРІНЬ ="):src.index("РЕСУРСИ = [")])
exec(src[src.index("def ресурс("):src.index("# --- регістр")])


def рядок(n):
    return f"""<v8:Type>xs:string</v8:Type>
					<v8:StringQualifiers>
						<v8:Length>{n}</v8:Length>
						<v8:AllowedLength>Variable</v8:AllowedLength>
					</v8:StringQualifiers>"""


РЕСУРСИ = [
    ("УсловияДоговора", "Условия договора", "Умови договору", рядок(1000)),
    ("СтатусВОтчете", "Статус в отчете", "Статус у звіті", рядок(10)),
]
т = читати(РЕГ)
перенос = "\r\n" if "\r\n" in т else "\n"
нові = "".join(ресурс(*р) for р in РЕСУРСИ if f"<Name>{р[0]}</Name>" not in т)
if нові:
    якір = "\t\t\t<Dimension uuid="
    assert т.count(якір) == 1
    т = т.replace(якір, нові.replace("\n", перенос) + якір, 1)
if "<Name>НеПоказыватьВОтчете</Name>" in т:
    м = т.index("<Name>НеПоказыватьВОтчете</Name>")
    початок = т.rfind("<Resource uuid=", 0, м)
    початок = т.rfind("\n", 0, початок) + 1
    кінець = т.index("</Resource>", м) + len("</Resource>")
    кінець = т.index("\n", кінець) + 1
    т = т[:початок] + т[кінець:]
писати(РЕГ, т)
minidom.parse(РЕГ)
т = читати(РЕГ)
assert т.count("<Resource uuid=") == 14 and "НеПоказыватьВОтчете" not in т, т.count("<Resource uuid=")
print("OK: регістр 14 ресурсів (+УсловияДоговора, +СтатусВОтчете, −НеПоказыватьВОтчете)")
