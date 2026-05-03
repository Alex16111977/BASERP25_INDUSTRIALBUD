"""Abstract Extractor."""
from __future__ import annotations

import abc


class Extractor(abc.ABC):
    """An Extractor pulls source rows out of 1С into list[dict] of SQL-typed values."""

    def __init__(self, config: dict) -> None:
        self.config = config

    @abc.abstractmethod
    def extract(self) -> list[dict]:
        """Return rows as a list of dicts. Keys are usually raw SQL column names
        (e.g. _IDRRef, _Fld55982). Transformers handle renaming and type fixups."""
