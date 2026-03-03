"""
Phone number normalisation.

Fixes the double-space bug present in the old Convert2.py where chunking
was applied to the already-prefixed "+91 XXXXXXXXXX" string, producing
"+91  XXXX XXXX XX" (extra space after country code).

Rules:
- Strip all non-digit characters.
- If exactly 10 digits and default country code is "+91", assume Indian number.
- If starts with 91 and is 12 digits, treat as Indian (strip leading 91).
- Any other length: prepend the country code digit-strip if not present.
- Format Indian numbers as "+91 XXXXX XXXXX" (5+5).
- Format other numbers as "+CC XXXXXXXXXX" (no internal spaces).
"""
from __future__ import annotations

import re


class PhoneNormalizer:
    def __init__(self, default_country_code: str = "+91"):
        # Store digits only (e.g. "91" from "+91")
        cc = re.sub(r"\D", "", default_country_code)
        self.country_code_digits: str = cc
        self.country_code: str = f"+{cc}"

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def normalize(self, raw: str) -> str:
        """
        Return a normalised phone string, e.g. "+91 98765 43210".
        Returns an empty string if the input is not a valid phone number.
        """
        if not raw:
            return ""

        digits = re.sub(r"\D", "", raw)

        if not digits:
            return ""

        # Already has country code prepended
        if raw.strip().startswith("+"):
            cc_digits = re.sub(r"\D", "", raw.split()[0] if " " in raw.strip() else raw)
            # Strip the country code part and normalise from national number
            if digits.startswith(self.country_code_digits):
                national = digits[len(self.country_code_digits):]
                return self._format(national)
            # Unknown country code – return as-is but clean
            return f"+{digits}"

        # 10-digit national number (no country code)
        if len(digits) == 10:
            return self._format(digits)

        # 11+ digits: might include country code
        if digits.startswith(self.country_code_digits):
            national = digits[len(self.country_code_digits):]
            if len(national) == 10:
                return self._format(national)

        # Fallback: just prepend country code
        if len(digits) >= 7:
            return self._format(digits)

        return ""

    def is_valid(self, raw: str) -> bool:
        """Returns True if the raw string contains a parseable phone number."""
        return bool(self.normalize(raw))

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _format(self, national_digits: str) -> str:
        """Format 10-digit national number as "+CC XXXXX XXXXX" (Indian style)
        or "+CC XXXXXXXXXX" for non-10-digit nationals."""
        if len(national_digits) == 10 and self.country_code_digits == "91":
            # Indian: +91 XXXXX XXXXX
            return f"{self.country_code} {national_digits[:5]} {national_digits[5:]}"
        # Generic: just attach country code with single space
        return f"{self.country_code} {national_digits}"
