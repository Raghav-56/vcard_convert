"""
Abstract base class for all parsers.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Union
from pathlib import Path

from vcard_convert.core.models import ParseResult


class AbstractParser(ABC):
    """
    All parsers take either a file path or a raw string and return a ParseResult.
    """

    @abstractmethod
    def parse_file(self, path: Union[str, Path]) -> ParseResult:
        """Parse a file on disk."""

    @abstractmethod
    def parse_text(self, text: str, source_label: str = "<text>") -> ParseResult:
        """Parse raw text content (e.g. pasted in the web UI)."""

    # Convenience: auto-dispatch based on whether argument is a path or text
    def parse(self, source: Union[str, Path]) -> ParseResult:
        p = Path(source)
        if p.exists() and p.is_file():
            return self.parse_file(p)
        return self.parse_text(str(source))
