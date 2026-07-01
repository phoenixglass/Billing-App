"""
Helper functions to:
- parse MMDDYYYY date tokens into weekday
- determine whether a given service is non-billable for a given weekday

Weekday mapping: 0=Monday, 1=Tuesday, ..., 6=Sunday

Weekly billing schedule:
- Monday (0):    Professional + Utox     (Programming + e-care non-billable)
- Tuesday (1):   E-care + Programming    (Professional + Utox non-billable)
- Wednesday (2): Professional + Utox     (Programming + e-care non-billable)
- Thursday (3):  Programming + Professional + Utox (e-care non-billable)
- Friday (4):    Programming + Professional + Utox (e-care non-billable)
- Saturday/Sunday (5,6): all services billable except e-care

Service categories:
- Programming: Detox, Residential, Partial Hospitalization (PHP), IOP
- Professional: all other services (including Acupuncture)
- Utox: drug screen services (matched by drug-screen keywords)

E-care notes:
- e-care is billable only on Tuesdays (non-billable on other days).
- Matches e-care variants: "e-care", "e care", "ecare", "extended care" (case-insensitive).

php_on_monday=True exempts Partial Hospitalization from the Monday restriction.

self_pay=True treats every service as billable every day, except e-care which
still only bills on Tuesdays.
"""
from datetime import datetime
from typing import Tuple


DRUG_SCREEN_KEYWORDS = ("drug screen", "utox", "urine tox", "drug test", "uds")


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


def _is_ecare(service: str) -> bool:
    """Return True if the service text refers to e-care (many possible variants)."""
    s = service.lower()
    return any(token in s for token in ("e-care", "e care", "ecare", "extended care"))


def _is_drug_screen(service: str) -> bool:
    """Return True if the service text refers to a drug screen (Utox)."""
    s = service.lower()
    return any(kw in s for kw in DRUG_SCREEN_KEYWORDS)


def _is_programming_service(service: str) -> bool:
    """Return True for Programming services: Detox, Residential, PHP, IOP."""
    s = service.lower()
    return (
        "detox" in s
        or "residential" in s
        or "partial hospitalization" in s
        or "php" in s
        or "iop" in s
    )


def is_non_billable_service_for_weekday(
    service: str, weekday: int, php_on_monday: bool = False, self_pay: bool = False
) -> bool:
    """
    Return True if the given service (string) should be treated as non-billable
    on the specified weekday.

    - service: original service string (case-insensitive checks will be used)
    - weekday: 0=Monday .. 6=Sunday
    - php_on_monday: when True, Partial Hospitalization is billable on Mondays
    - self_pay: when True, every service is billable every day except e-care,
      which still only bills on Tuesdays
    """
    s = (service or "").lower().strip()

    # e-care is only billed on Tuesdays (applies to self-pay too)
    if _is_ecare(s):
        return weekday != 1

    # Self Pay: every other service is billable every day
    if self_pay:
        return False

    # Drug screens (Utox) follow Professional days: billable Mon/Wed/Thu/Fri
    # plus weekends; non-billable only on Tuesday.
    if _is_drug_screen(s):
        return weekday == 1

    is_programming = _is_programming_service(s)

    # Monday and Wednesday: Professional + Utox only (Programming non-billable)
    if weekday in (0, 2):
        if not is_programming:
            return False
        # php_on_monday exempts Partial Hospitalization on Mondays
        if weekday == 0 and php_on_monday and (
            "partial hospitalization" in s or "php" in s
        ):
            return False
        return True

    # Tuesday: E-care + Programming only (Professional + Utox non-billable)
    if weekday == 1:
        return not is_programming

    # Thursday and Friday: Programming + Professional + Utox all billable
    if weekday in (3, 4):
        return False

    # Saturday and Sunday: everything billable (e-care already handled above)
    return False
