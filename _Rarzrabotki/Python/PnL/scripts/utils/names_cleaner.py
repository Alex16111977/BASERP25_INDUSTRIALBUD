"""Normalize names for fuzzy matching."""
import re

_REPLACE_PUNCT = re.compile(r"[.,;:\"'`«»()\[\]№\-–—/\\]+")
_SPACES = re.compile(r"\s+")


def normalize(name: str) -> str:
    if not name:
        return ""
    s = str(name)
    s = _REPLACE_PUNCT.sub(" ", s)
    s = _SPACES.sub(" ", s).strip()
    return s.casefold()
