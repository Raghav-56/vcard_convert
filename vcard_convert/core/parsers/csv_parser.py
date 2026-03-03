"""
CSV → Contacts parser.

Handles:
- Standard CSV with Name/Phone columns (any order, case-insensitive).
- Google Contacts Takeout CSV (Given Name, Family Name, Phone 1 - Value, …).
- UTF-8-BOM files exported from Excel/Google Sheets.
- Multi-contact rows where the Name cell contains comma-separated names.
"""
from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

from vcard_convert.core.models import Contact, ParseError, ParseResult, SourceFormat
from vcard_convert.core.parsers.base import AbstractParser
from vcard_convert.core.phone import PhoneNormalizer

# ── Column name aliases ──────────────────────────────────────────────────────

_STANDARD_NAME_COLS = {"name", "full name", "contact name", "fullname"}
_STANDARD_PHONE_COLS = {"phone", "phone number", "mobile", "cell", "tel", "number"}
_STANDARD_EMAIL_COLS = {"email", "e-mail", "email address"}
_STANDARD_ORG_COLS = {"organization", "org", "company"}
_STANDARD_TITLE_COLS = {"title", "job title", "position"}
_STANDARD_ADDR_COLS = {"address", "addr"}
_STANDARD_NOTES_COLS = {"notes", "note", "remarks"}

# Google Contacts Takeout column patterns
_GOOGLE_GIVEN_NAME = "given name"
_GOOGLE_FAMILY_NAME = "family name"
_GOOGLE_PHONE_RE = re.compile(r"^phone \d+ - value$", re.IGNORECASE)
_GOOGLE_EMAIL_RE = re.compile(r"^e-mail \d+ - value$", re.IGNORECASE)
_GOOGLE_ORG_RE = re.compile(r"^organization \d+ - name$", re.IGNORECASE)


def _norm_col(c: str) -> str:
    return c.strip().lower()


class CSVParser(AbstractParser):
    def __init__(
        self,
        country_code: str = "+91",
        force_google_mode: bool = False,
    ):
        self.normalizer = PhoneNormalizer(country_code)
        self.force_google_mode = force_google_mode

    # ── Public ────────────────────────────────────────────────────────────────

    def parse_file(self, path: Union[str, Path]) -> ParseResult:
        path = Path(path)
        text = path.read_text(encoding="utf-8-sig")  # handles BOM
        return self.parse_text(text, source_label=str(path))

    def parse_text(self, text: str, source_label: str = "<csv>") -> ParseResult:
        reader = csv.DictReader(text.splitlines())
        if reader.fieldnames is None:
            return ParseResult(
                source_format=SourceFormat.CSV,
                source_path=source_label,
                errors=[ParseError(raw=text[:80], reason="File appears to be empty or not valid CSV")],
            )

        col_map = self._map_columns(list(reader.fieldnames))
        is_google = self.force_google_mode or col_map.get("_is_google", False)
        fmt = SourceFormat.GOOGLE_CSV if is_google else SourceFormat.CSV

        contacts: List[Contact] = []
        errors: List[ParseError] = []

        for i, row in enumerate(reader, start=2):  # row 1 is header
            parsed = self._parse_row(row, col_map, line_number=i)
            contacts.extend(parsed[0])
            errors.extend(parsed[1])

        return ParseResult(
            contacts=contacts,
            errors=errors,
            source_format=fmt,
            source_path=source_label,
        )

    # ── Column detection ──────────────────────────────────────────────────────

    def _map_columns(self, fieldnames: List[str]) -> Dict[str, str]:
        """
        Return a mapping of semantic role → actual column name.
        Also sets "_is_google": True if Google Contacts header detected.
        """
        normed = {_norm_col(f): f for f in fieldnames}
        mapping: Dict[str, str] = {}

        _google_given_name_col = normed.get(_GOOGLE_GIVEN_NAME)
        is_google = bool(_google_given_name_col)

        if is_google:
            mapping["_is_google"] = True
            mapping["first_name"] = normed[_GOOGLE_GIVEN_NAME]
            if fn := normed.get(_GOOGLE_FAMILY_NAME):
                mapping["last_name"] = fn
            # phone / email / org: find all matching columns
            phone_cols = [orig for n, orig in normed.items() if _GOOGLE_PHONE_RE.match(n)]
            email_cols = [orig for n, orig in normed.items() if _GOOGLE_EMAIL_RE.match(n)]
            org_cols = [orig for n, orig in normed.items() if _GOOGLE_ORG_RE.match(n)]
            if phone_cols:
                mapping["phones"] = phone_cols  # list of columns
            if email_cols:
                mapping["emails"] = email_cols
            if org_cols:
                mapping["organization"] = org_cols[0]
        else:
            # Standard mode: find best-match column per field
            for role, aliases in [
                ("name", _STANDARD_NAME_COLS),
                ("phone", _STANDARD_PHONE_COLS),
                ("email", _STANDARD_EMAIL_COLS),
                ("organization", _STANDARD_ORG_COLS),
                ("title", _STANDARD_TITLE_COLS),
                ("address", _STANDARD_ADDR_COLS),
                ("notes", _STANDARD_NOTES_COLS),
            ]:
                for alias in aliases:
                    if alias in normed:
                        mapping[role] = normed[alias]
                        break

        return mapping

    # ── Row parsing ───────────────────────────────────────────────────────────

    def _parse_row(
        self, row: Dict[str, str], col_map: Dict[str, str], line_number: int
    ) -> Tuple[List[Contact], List[ParseError]]:

        is_google = "_is_google" in col_map

        if is_google:
            return self._parse_google_row(row, col_map, line_number)

        # Standard mode
        raw_name = row.get(col_map.get("name", ""), "").strip()
        raw_phone = row.get(col_map.get("phone", ""), "").strip()

        if not raw_name and not raw_phone:
            return [], []  # blank row – skip silently

        if not raw_name:
            return [], [ParseError(raw=str(row), reason="Missing name", line_number=line_number)]

        if not raw_phone:
            return [], [ParseError(raw=str(row), reason="Missing phone number", line_number=line_number)]

        # Multi-contact row detection: "Alice, Bob" with "0987..., 1234..." 
        name_parts = [n.strip() for n in raw_name.split(",") if n.strip()]
        phone_parts = [p.strip() for p in re.split(r"[,/]", raw_phone) if p.strip()]

        if len(name_parts) > 1 and len(phone_parts) > 1:
            return self._parse_multi(name_parts, phone_parts, row, col_map, line_number)

        # Single contact
        contact, err = self._build_contact(raw_name, raw_phone, row, col_map, line_number)
        if err:
            return [], [err]
        return [contact], []

    def _parse_google_row(
        self, row: Dict[str, str], col_map: Dict[str, str], line_number: int
    ) -> Tuple[List[Contact], List[ParseError]]:
        first = row.get(col_map.get("first_name", ""), "").strip()
        last = row.get(col_map.get("last_name", ""), "").strip()

        phone_cols = col_map.get("phones", [])
        phones = [
            self.normalizer.normalize(row.get(c, ""))
            for c in phone_cols
            if row.get(c, "").strip()
        ]
        phones = [p for p in phones if p]

        if not first and not last:
            return [], []
        if not phones:
            return [], [ParseError(raw=str(row), reason="Missing phone number", line_number=line_number)]

        email_cols = col_map.get("emails", [])
        emails = [row.get(c, "").strip() for c in email_cols if row.get(c, "").strip()]

        c = Contact(
            first_name=first,
            last_name=last,
            phones=phones,
            emails=emails,
            organization=row.get(col_map.get("organization", ""), "").strip(),
        )
        return [c], []

    def _parse_multi(
        self,
        names: List[str],
        phones: List[str],
        row: Dict[str, str],
        col_map: Dict[str, str],
        line_number: int,
    ) -> Tuple[List[Contact], List[ParseError]]:
        contacts: List[Contact] = []
        errors: List[ParseError] = []
        pairs = zip(names, phones)
        for name, phone in pairs:
            c, err = self._build_contact(name, phone, row, col_map, line_number)
            if err:
                errors.append(err)
            else:
                contacts.append(c)
        return contacts, errors

    def _build_contact(
        self, name: str, raw_phone: str, row: Dict[str, str], col_map: Dict[str, str], line_number: int
    ) -> Tuple[Optional[Contact], Optional[ParseError]]:
        if len(name) < 2:
            return None, ParseError(raw=name, reason=f"Name too short: {name!r}", line_number=line_number)

        phone = self.normalizer.normalize(raw_phone)
        if not phone:
            return None, ParseError(
                raw=raw_phone, reason=f"Invalid phone number: {raw_phone!r}", line_number=line_number
            )

        # Split name into first / last (simple heuristic)
        parts = name.split()
        first = parts[0] if parts else name
        last = " ".join(parts[1:]) if len(parts) > 1 else ""

        contact = Contact(
            first_name=first,
            last_name=last,
            phones=[phone],
            emails=[row.get(col_map.get("email", ""), "").strip()] if col_map.get("email") else [],
            organization=row.get(col_map.get("organization", ""), "").strip() if col_map.get("organization") else "",
            title=row.get(col_map.get("title", ""), "").strip() if col_map.get("title") else "",
            address=row.get(col_map.get("address", ""), "").strip() if col_map.get("address") else "",
            notes=row.get(col_map.get("notes", ""), "").strip() if col_map.get("notes") else "",
            raw_line=str(row),
        )
        # Drop empty email entries
        contact.emails = [e for e in contact.emails if e]
        return contact, None
