"""
Helper functions to:
- parse MMDDYYYY date tokens into weekday
- determine whether a given service is non-billable for a given weekday

Weekday mapping: 0=Monday, 1=Tuesday, ..., 6=Sunday

Rules implemented:
- Tuesday (1): everything billed (including e-care).
- Monday (0): non-billable: partial hospitalization, residential, detox, and e-care.
- Wednesday-Friday (2,3,4): all services billed except e-care.
- e-care is billable only on Tuesdays (non-billable on other days).
- Matches e-care variants: "e-care", "e care", "ecare" (case-insensitive).
"""
from datetime import datetime
from typing import Tuple


def parse_weekday_from_token(date_token: str) -> Tuple[int, bool]:
    """
    Parse an MMDDYYYY date_token and return (weekday, did_fallback).
    - weekday: integer 0..6
    - did_fallback: True if parsing failed and we fell back to today's weekday
    """
    if not date_token:
        return datetime.now().weekday(), True
    try:
        dt = datetime.strptime(date_token, "%m%d%Y")
        return dt.weekday(), False
    except Exception:
        return datetime.now().weekday(), True


def _is_ecare(service_lower: str) -> bool:
    """Return True if the service text refers to e-care (many possible variants)."""
    return any(token in service_lower for token in ("e-care", "e care", "ecare"))


def is_non_billable_service_for_weekday(service: str, weekday: int) -> bool:
    """
    Return True if the given service (string) should be treated as non-billable
    on the specified weekday.

    - service: original service string (case-insensitive checks will be used)
    - weekday: 0=Monday .. 6=Sunday
    """
    s = (service or "").lower().strip()

    # e-care is only billed on Tuesdays
    if _is_ecare(s):
        return weekday != 1  # non-billable unless Tuesday

    # Monday: exclude partial hospitalization, residential, detox
    if weekday == 0:
        if any(x in s for x in ("partial hospitalization", "residential", "detox")):
            return True
        return False

    # Tuesday: everything billable (already covered e-care above)
    if weekday == 1:
        return False

    # Wed-Fri and weekends: only e-care is non-billable (handled above)
    return False
