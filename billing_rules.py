"""
Helper functions to:
- parse MMDDYYYY date tokens into weekday
- determine whether a given service is non-billable for a given weekday
- classify a "Claim Type" value as Professional (CMS-1500)

Weekday mapping: 0=Monday, 1=Tuesday, ..., 6=Sunday

Daily billing schedule:
- Professional services (Claim Type "CMS-1500") bill every day of the week.
- Programming (Detox, Residential, Partial Hospitalization/PHP, IOP) bills
  Tuesday, Thursday, Friday, and weekends; non-billable Monday and
  Wednesday.
- E-care bills on Tuesdays only, regardless of Claim Type.
- Self Pay: every service bills every day, with no exceptions (including
  e-care).

Rosanna/Jasmine split (applies to GROUPFLD2 == "Insurance" rows only; no
Self Pay ever goes to either of them):
- The "professional pool" is every Insurance row whose Claim Type is
  CMS-1500, sorted alphabetically by Client. Rosanna receives the first
  ROSANNA_PROFESSIONAL_CAP[weekday] rows of that sorted pool (0 if the
  weekday isn't in the map, i.e. Wednesday and weekends); the remainder of
  the pool goes to Jasmine.
- Jasmine also receives every billable Programming/e-care row.
- IOP (including Telemed IOP) always goes to Jasmine when billable that
  weekday, bypassing the professional pool/Rosanna split entirely, even if
  Claim Type is CMS-1500.
- GROUPFLD2 values other than "Insurance" or "Self Pay" never reach
  Rosanna or Jasmine.
"""
from datetime import datetime
from typing import Tuple


DRUG_SCREEN_KEYWORDS = ("drug screen", "utox", "urine tox", "drug test", "uds")

# Rosanna's professional-service row cap by weekday (0=Monday..6=Sunday).
# Wednesday (2) and weekends (5, 6) are intentionally absent: Rosanna
# receives no rows those days and Jasmine gets the entire professional pool.
ROSANNA_PROFESSIONAL_CAP = {0: 300, 1: 300, 3: 125, 4: 125}


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


def is_iop_service(service: str) -> bool:
    """Return True for any IOP service, including Telemed IOP (case-insensitive)."""
    return "iop" in (service or "").lower()


def is_professional_claim_type(claim_type: str) -> bool:
    """Return True if the Claim Type column value is CMS-1500 (Professional)."""
    return (claim_type or "").strip().upper() == "CMS-1500"


def is_non_billable_service_for_weekday(
    service: str, weekday: int, is_professional: bool = False, self_pay: bool = False
) -> bool:
    """
    Return True if the given service (string) should be treated as non-billable
    on the specified weekday.

    - service: original service string (case-insensitive checks will be used)
    - weekday: 0=Monday .. 6=Sunday
    - is_professional: True when the row's Claim Type is CMS-1500. Professional
      rows bill every day of the week.
    - self_pay: when True, every service is billable every day, with no
      exceptions (including e-care).
    """
    s = (service or "").lower().strip()

    # Self Pay bills every service every day, no exceptions.
    if self_pay:
        return False

    # Professional (CMS-1500) services bill every day of the week.
    if is_professional:
        return False

    # e-care is only billed on Tuesdays.
    if _is_ecare(s):
        return weekday != 1

    # Programming (Detox, Residential, PHP, IOP) bills Tue/Thu/Fri + weekends.
    if _is_programming_service(s):
        return weekday not in (1, 3, 4, 5, 6)

    # Anything else is neither Professional nor Programming nor e-care, and
    # isn't addressed by the daily schedule, so it is not billable.
    return True
