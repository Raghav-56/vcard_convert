"""
vCard (.vcf) parser — reads vCard 2.1 and 3.0 files into Contact objects.
Used for the merge / reformat workflow.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Union

from vcard_convert.core.models import Contact, ParseError, ParseResult, SourceFormat
from vcard_convert.core.parsers.base import AbstractParser
from vcard_convert.core.phone import PhoneNormalizer


def _unescape(value: str) -> str:
    """Remove vCard 3.0 property value escaping."""
    return value.replace("\\n", "\n").replace("\\,", ",").replace("\\;", ";")


class VCFParser(AbstractParser):
    def __init__(self, country_code: str = "+91"):
        self.normalizer = PhoneNormalizer(country_code)

    # ── Public ────────────────────────────────────────────────────────────────

    def parse_file(self, path: Union[str, Path]) -> ParseResult:
        path = Path(path)
        text = path.read_text(encoding="utf-8-sig")
        return self.parse_text(text, source_label=str(path))

    def parse_text(self, text: str, source_label: str = "<vcf>") -> ParseResult:
        blocks = self._split_blocks(text)
        contacts: List[Contact] = []
        errors: List[ParseError] = []

        for i, block in enumerate(blocks, start=1):
            result = self._parse_block(block, block_index=i)
            if isinstance(result, Contact):
                contacts.append(result)
            else:
                errors.append(result)

        return ParseResult(
            contacts=contacts,
            errors=errors,
            source_format=SourceFormat.VCF,
            source_path=source_label,
        )

    # ── Internals ─────────────────────────────────────────────────────────────

    def _split_blocks(self, text: str) -> List[str]:
        """Split on BEGIN:VCARD / END:VCARD pairs, handling CRLF/LF."""
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        blocks: List[str] = []
        current: List[str] = []
        in_block = False
        for line in text.splitlines():
            if line.upper() == "BEGIN:VCARD":
                in_block = True
                current = [line]
            elif line.upper() == "END:VCARD":
                current.append(line)
                blocks.append("\n".join(current))
                in_block = False
                current = []
            elif in_block:
                current.append(line)
        return blocks

    def _parse_block(self, block: str, block_index: int) -> Union[Contact, ParseError]:
        props: Dict[str, List[str]] = {}  # prop name → list of values

        for line in block.splitlines():
            if line.upper() in ("BEGIN:VCARD", "END:VCARD"):
                continue
            if ":" not in line:
                continue
            prop_part, _, value = line.partition(":")
            # prop_part can be "TEL;TYPE=CELL" → key = "TEL"
            prop_name = prop_part.split(";")[0].upper()
            props.setdefault(prop_name, []).append(_unescape(value.strip()))

        # Full name
        fn_values = props.get("FN", [])
        n_values = props.get("N", [])

        full_name = fn_values[0] if fn_values else ""
        first_name, last_name = "", ""

        if n_values:
            # N:Lastname;Firstname;Middle;Prefix;Suffix
            n_parts = n_values[0].split(";")
            last_name = n_parts[0].strip() if len(n_parts) > 0 else ""
            first_name = n_parts[1].strip() if len(n_parts) > 1 else ""

        if not first_name and not last_name and full_name:
            parts = full_name.split()
            first_name = parts[0] if parts else full_name
            last_name = " ".join(parts[1:]) if len(parts) > 1 else ""

        if not first_name and not last_name:
            return ParseError(
                raw=block[:80],
                reason="No name found in vCard block",
                line_number=block_index,
            )

        # Phones
        raw_phones = props.get("TEL", [])
        phones = [self.normalizer.normalize(p) for p in raw_phones]
        phones = [p for p in phones if p]

        if not phones:
            return ParseError(
                raw=block[:80],
                reason=f"No valid phone in vCard for '{full_name or first_name}'",
                line_number=block_index,
            )

        # Optional fields
        emails = [e.strip() for e in props.get("EMAIL", []) if e.strip()]
        org_values = props.get("ORG", [])
        org = org_values[0].split(";")[0].strip() if org_values else ""
        title = props.get("TITLE", [""])[0].strip()
        adr_values = props.get("ADR", [])
        address = "; ".join(adr_values).strip() if adr_values else ""
        note_values = props.get("NOTE", [])
        notes = " ".join(note_values).strip()

        return Contact(
            first_name=first_name,
            last_name=last_name,
            phones=phones,
            emails=emails,
            organization=org,
            title=title,
            address=address,
            notes=notes,
            raw_line=block,
        )
