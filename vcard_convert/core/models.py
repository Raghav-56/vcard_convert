"""
Core data models for the contact conversion tool.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


class SourceFormat(str, Enum):
    CSV = "csv"
    GOOGLE_CSV = "google-csv"
    VCF = "vcf"
    TEXT = "text"
    UNKNOWN = "unknown"


class OutputFormat(str, Enum):
    VCF = "vcf"


class VCardVersion(str, Enum):
    V2_1 = "2.1"
    V3_0 = "3.0"


@dataclass
class Contact:
    """Single contact with all possible fields."""

    first_name: str = ""
    last_name: str = ""
    phones: List[str] = field(default_factory=list)
    emails: List[str] = field(default_factory=list)
    organization: str = ""
    title: str = ""
    address: str = ""
    notes: str = ""
    # Internal bookkeeping
    raw_line: str = ""
    confidence: float = 1.0  # 0.0–1.0; < 1.0 means the text parser was unsure

    @property
    def full_name(self) -> str:
        parts = [self.first_name, self.last_name]
        return " ".join(p for p in parts if p).strip()

    @property
    def is_valid(self) -> bool:
        return bool(self.full_name) and bool(self.phones)

    def __repr__(self) -> str:  # pragma: no cover
        return f"Contact({self.full_name!r}, phones={self.phones})"


@dataclass
class ParseError:
    """A row / line that could not be turned into a Contact."""

    raw: str
    reason: str
    line_number: Optional[int] = None


@dataclass
class ParseResult:
    """Everything returned by a parser."""

    contacts: List[Contact] = field(default_factory=list)
    errors: List[ParseError] = field(default_factory=list)
    source_format: SourceFormat = SourceFormat.UNKNOWN
    source_path: str = ""

    @property
    def total(self) -> int:
        return len(self.contacts) + len(self.errors)

    @property
    def success_rate(self) -> float:
        if self.total == 0:
            return 0.0
        return len(self.contacts) / self.total
