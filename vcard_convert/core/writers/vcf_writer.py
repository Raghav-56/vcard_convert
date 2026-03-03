"""
vCard writer — serialises Contact objects to vCard 3.0 (or 2.1) format.

Fixes in comparison to the old code:
- No double-space in phone numbers (PhoneNormalizer handles formatting).
- Correct TEL;TYPE=CELL syntax for vCard 3.0.
- Proper N:Lastname;Firstname;;; field.
- All optional fields (EMAIL, ORG, TITLE, ADR, NOTE) are included when present.
- Optional duplicate detection by normalised phone number.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Set, Union

from vcard_convert.core.models import Contact, VCardVersion
from vcard_convert.core.writers.base import AbstractWriter


def _escape(value: str) -> str:
    """vCard 3.0 property value escaping."""
    return value.replace("\\", "\\\\").replace(",", "\\,").replace(";", "\\;").replace("\n", "\\n")


class VCFWriter(AbstractWriter):
    def __init__(
        self,
        version: VCardVersion = VCardVersion.V3_0,
        dedupe: bool = False,
    ):
        self.version = version
        self.dedupe = dedupe

    # ── Public ────────────────────────────────────────────────────────────────

    def write_file(self, contacts: List[Contact], path: Union[str, Path]) -> None:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        text = self.write_text(contacts)
        path.write_text(text, encoding="utf-8")

    def write_text(self, contacts: List[Contact]) -> str:
        seen_phones: Set[str] = set()
        blocks: List[str] = []

        for contact in contacts:
            if self.dedupe:
                # Deduplicate on normalised (digits-only) phone
                key = re.sub(r"\D", "", contact.phones[0]) if contact.phones else ""
                if key and key in seen_phones:
                    continue
                if key:
                    seen_phones.add(key)

            block = (
                self._make_vcard_3(contact)
                if self.version == VCardVersion.V3_0
                else self._make_vcard_21(contact)
            )
            blocks.append(block)

        return "\n".join(blocks) + ("\n" if blocks else "")

    # ── vCard 3.0 ─────────────────────────────────────────────────────────────

    def _make_vcard_3(self, c: Contact) -> str:
        lines: List[str] = ["BEGIN:VCARD", "VERSION:3.0"]

        fn = _escape(c.full_name)
        lines.append(f"FN:{fn}")

        last = _escape(c.last_name)
        first = _escape(c.first_name)
        lines.append(f"N:{last};{first};;;")

        for phone in c.phones:
            lines.append(f"TEL;TYPE=CELL:{phone}")

        for email in c.emails:
            if email:
                lines.append(f"EMAIL;TYPE=INTERNET:{_escape(email)}")

        if c.organization:
            lines.append(f"ORG:{_escape(c.organization)}")

        if c.title:
            lines.append(f"TITLE:{_escape(c.title)}")

        if c.address:
            lines.append(f"ADR;TYPE=HOME:;;{_escape(c.address)};;;;")

        if c.notes:
            lines.append(f"NOTE:{_escape(c.notes)}")

        lines.append("END:VCARD")
        return "\n".join(lines)

    # ── vCard 2.1 ─────────────────────────────────────────────────────────────

    def _make_vcard_21(self, c: Contact) -> str:
        lines: List[str] = ["BEGIN:VCARD", "VERSION:2.1"]

        fn = c.full_name
        lines.append(f"FN:{fn}")
        lines.append(f"N;CHARSET=UTF-8:{c.last_name};{c.first_name};;;")

        for phone in c.phones:
            lines.append(f"TEL;CELL;PREF:{phone}")

        for email in c.emails:
            if email:
                lines.append(f"EMAIL;INTERNET:{email}")

        if c.organization:
            lines.append(f"ORG:{c.organization}")

        if c.title:
            lines.append(f"TITLE:{c.title}")

        if c.notes:
            lines.append(f"NOTE:{c.notes}")

        lines.append("END:VCARD")
        return "\n".join(lines)
