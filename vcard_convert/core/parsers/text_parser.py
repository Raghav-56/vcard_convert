"""
Free-text → Contacts parser.

Handles three common paste patterns without any heavy NLP dependencies:

1. WhatsApp group member export format:
      Raj Kumar ~ +91 98765 43210
      Priya     ~ 87654 32109

2. Copied spreadsheet rows (tab or comma-separated):
      Raj Kumar\t9876543210
      Priya, 8765432109

3. Loose prose:
      "Call Raj at 9876543210 or Priya on 8765432109"

Low-confidence extractions (confidence < 1.0) are flagged so the user can
review them before export.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Tuple, Union

from vcard_convert.core.models import Contact, ParseError, ParseResult, SourceFormat
from vcard_convert.core.parsers.base import AbstractParser
from vcard_convert.core.phone import PhoneNormalizer

# Regex that finds a plausible phone number within a larger string.
# Matches: +91 98765 43210, 987-654-3210, (987) 654 3210, 9876543210, etc.
_PHONE_RE = re.compile(
    r"""
    (?<!\d)                        # not preceded by a digit
    (\+?[\d][\d\s\-().]{7,17}\d)  # 8–18 char phone-like string
    (?!\d)                         # not followed by a digit
    """,
    re.VERBOSE,
)

# WhatsApp style separator: "Name ~ number" or "Name - number"
_WHATSAPP_SEP_RE = re.compile(r"\s*[~\-–—:]\s*")

# Prose trigger words that hint the following token is a name
_PROSE_NAME_RE = re.compile(
    r"(?:call|contact|text|message|reach|ping)\s+([A-Za-z][A-Za-z\s]{1,30}?)\s+(?:at|on|@)",
    re.IGNORECASE,
)


class FreeTextParser(AbstractParser):
    def __init__(self, country_code: str = "+91"):
        self.normalizer = PhoneNormalizer(country_code)

    # ── Public ────────────────────────────────────────────────────────────────

    def parse_file(self, path: Union[str, Path]) -> ParseResult:
        text = Path(path).read_text(encoding="utf-8-sig")
        return self.parse_text(text, source_label=str(path))

    def parse_text(self, text: str, source_label: str = "<text>") -> ParseResult:
        contacts: List[Contact] = []
        errors: List[ParseError] = []

        lines = [l.strip() for l in text.replace("\r\n", "\n").replace("\r", "\n").splitlines()]
        i = 0
        while i < len(lines):
            line = lines[i]
            if not line:
                i += 1
                continue

            c, err = self._parse_line(line, line_number=i + 1)
            if c:
                contacts.append(c)
            elif err:
                # Try multi-line: maybe name is on line i and phone is on i+1
                if i + 1 < len(lines) and lines[i + 1]:
                    combined = line + " " + lines[i + 1]
                    c2, err2 = self._parse_line(combined, line_number=i + 1)
                    if c2:
                        c2.confidence = 0.7  # lower confidence for combined-line match
                        contacts.append(c2)
                        i += 2
                        continue
                errors.append(err)
            i += 1

        # Prose pass: scan the whole text for "call X at Y" patterns
        for match in _PROSE_NAME_RE.finditer(text):
            name_raw = match.group(1).strip()
            # Look for a phone number immediately after the "at/on"
            after = text[match.end():]
            pm = _PHONE_RE.match(after)
            if pm:
                phone = self.normalizer.normalize(pm.group(1))
                if phone:
                    parts = name_raw.split()
                    c = Contact(
                        first_name=parts[0],
                        last_name=" ".join(parts[1:]) if len(parts) > 1 else "",
                        phones=[phone],
                        confidence=0.8,
                        raw_line=match.group(0),
                    )
                    # Avoid duplicates (phone already extracted by line parser)
                    if not any(phone in existing.phones for existing in contacts):
                        contacts.append(c)

        return ParseResult(
            contacts=contacts,
            errors=errors,
            source_format=SourceFormat.TEXT,
            source_path=source_label,
        )

    # ── Internals ─────────────────────────────────────────────────────────────

    def _parse_line(
        self, line: str, line_number: int
    ) -> Tuple[Optional[Contact], Optional[ParseError]]:
        # ── Strategy 1: WhatsApp "Name ~ number" ──────────────────────────
        if _WHATSAPP_SEP_RE.search(line):
            parts = _WHATSAPP_SEP_RE.split(line, maxsplit=1)
            if len(parts) == 2:
                name_raw, phone_raw = parts[0].strip(), parts[1].strip()
                phone = self.normalizer.normalize(phone_raw)
                if phone and len(name_raw) >= 2:
                    return self._make_contact(name_raw, [phone], line), None

        # ── Strategy 2: tab-separated ─────────────────────────────────────
        if "\t" in line:
            tab_parts = [p.strip() for p in line.split("\t") if p.strip()]
            return self._try_name_phone_parts(tab_parts, line, line_number)

        # ── Strategy 3: all phone numbers in the line ─────────────────────
        phone_matches = list(_PHONE_RE.finditer(line))
        if not phone_matches:
            return None, ParseError(
                raw=line, reason="No phone number found", line_number=line_number
            )

        phones = [self.normalizer.normalize(m.group(1)) for m in phone_matches]
        phones = [p for p in phones if p]

        # Remove phone spans from line to derive name
        name_raw = line
        for m in reversed(phone_matches):
            name_raw = name_raw[: m.start()] + name_raw[m.end():]
        name_raw = re.sub(r"[,;/\s]+", " ", name_raw).strip().strip(",;-~").strip()

        if len(name_raw) >= 2 and phones:
            confidence = 1.0 if len(name_raw) > 3 else 0.6
            c = self._make_contact(name_raw, phones, line)
            c.confidence = confidence
            return c, None

        if phones:
            # Phone found but no name — create placeholder so user can review
            c = Contact(
                first_name="Unknown",
                phones=phones,
                confidence=0.3,
                raw_line=line,
            )
            return c, None

        return None, ParseError(raw=line, reason="Could not extract name or phone", line_number=line_number)

    def _try_name_phone_parts(
        self, parts: List[str], line: str, line_number: int
    ) -> Tuple[Optional[Contact], Optional[ParseError]]:
        """For tab/comma split parts, find which part is name and which is phone."""
        name_raw: Optional[str] = None
        phones: List[str] = []

        for part in parts:
            phone = self.normalizer.normalize(part)
            if phone:
                phones.append(phone)
            elif len(part) >= 2:
                name_raw = part

        if name_raw and phones:
            return self._make_contact(name_raw, phones, line), None
        if phones:
            c = Contact(first_name="Unknown", phones=phones, confidence=0.3, raw_line=line)
            return c, None
        return None, ParseError(raw=line, reason="Could not parse tab-separated row", line_number=line_number)

    @staticmethod
    def _make_contact(name_raw: str, phones: List[str], raw_line: str) -> Contact:
        parts = name_raw.split()
        return Contact(
            first_name=parts[0] if parts else name_raw,
            last_name=" ".join(parts[1:]) if len(parts) > 1 else "",
            phones=phones,
            raw_line=raw_line,
        )
