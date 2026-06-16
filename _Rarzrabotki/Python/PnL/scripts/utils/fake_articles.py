# -*- coding: utf-8 -*-
"""Реестр фейк-статей PL: match_fake(name) -> reason|None.

Источник правил — data/fake_articles.json (substrings / exact_names / regexes).
Используется excel_parser.classify_row для отсева служебных строк финансиста,
которые НЕ являются статьями PL. Пополняется правкой JSON, без изменения кода.
"""
import json
import re
from pathlib import Path

_JSON = Path(__file__).resolve().parents[2] / "data" / "fake_articles.json"
_CACHE = None


def _norm(s):
    return " ".join(str(s or "").split()).strip()


def _load():
    global _CACHE
    if _CACHE is None:
        d = json.loads(_JSON.read_text(encoding="utf-8"))
        subs = [(_norm(x["pattern"]).casefold(), x.get("reason", "")) for x in d.get("substrings", [])]
        exacts = {_norm(x["pattern"]).casefold(): x.get("reason", "") for x in d.get("exact_names", [])}
        regexes = []
        for x in d.get("regexes", []):
            flags = re.IGNORECASE if "i" in str(x.get("flags", "")).lower() else 0
            regexes.append((re.compile(x["pattern"], flags), x.get("reason", "")))
        _CACHE = (subs, exacts, regexes)
    return _CACHE


def match_fake(name):
    """Вернуть причину (str), если name — фейк-статья; иначе None."""
    s = _norm(name)
    if not s:
        return None
    sl = s.casefold()
    subs, exacts, regexes = _load()
    if sl in exacts:
        return exacts[sl]
    for pat, reason in subs:
        if pat and pat in sl:
            return reason
    for rx, reason in regexes:
        if rx.search(s):
            return reason
    return None


def is_fake(name):
    return match_fake(name) is not None
