"""
Abstract base class for all writers.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import List, Union

from vcard_convert.core.models import Contact


class AbstractWriter(ABC):
    @abstractmethod
    def write_file(self, contacts: List[Contact], path: Union[str, Path]) -> None:
        """Write contacts to a file on disk."""

    @abstractmethod
    def write_text(self, contacts: List[Contact]) -> str:
        """Serialise contacts to a string (for web UI download)."""
