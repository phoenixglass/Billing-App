"""
Helper functions to:
- parse MMDDYYYY date tokens into weekday
- determine whether a given service is non-billable for a given weekday
- classify a "Claim Type" value as Professional (CMS-1500 or UB-04)

Weekday mapping: 0=Monday, 1=Tuesday, ..., 6=Sunday

Daily billing schedule:
- Professional services (Claim Type "CMS-1500" or "UB-04" — UB-04 counts as
  Professional every day of the week) bill every day of the week.
- IOP (including Telemed IOP) bills every day of the week, with no
  exceptions.
- Programming (Detox, Residential) bills Tuesday, Thursday, Friday, and
  weekends; non-billable Monday and Wednesday.
- E-care bills on Tuesdays only, regardless of Claim Type.
- Self Pay: every service bills every day, with no exceptions (including
  e-care).
- PHP (Partial Hospitalization) is not part of the Programming schedule
  above: it always goes to Melissa (see is_php_service), and is billed
  only on Tuesdays as a real-world/manual matter, not something this
  module's weekday helpers gate.

Rosanna/Jasmine split (applies to GROUPFLD2 == "Insurance" rows only; no
Self Pay ever goes to either of them, and PHP rows never reach them since
PHP always goes to Melissa):
- The "professional pool" is every Insurance row whose Claim Type is
  CMS-1500 or UB-04, sorted alphabetically by Client. Monday through
  Friday, Rosanna receives the first ROSANNA_PROFESSIONAL_CAP[weekday]
  rows (150) of that sorted pool; anything past her share of the pool
  goes to Jasmine. Rosanna caps no rows on weekends, so the entire pool
  goes to Jasmine those days.
- Jasmine also receives every billable Programming/e-care row, plus any
  other billable Insurance row whose Claim Type is not CMS-1500/UB-04
  (i.e. institutional/837I), unless it's PHP (always Melissa's).
- IOP (including Telemed IOP) always goes to Jasmine, every day of the
  week, bypassing the professional pool/Rosanna split entirely, even if
  Claim Type is CMS-1500 or UB-04.
- GROUPFLD2 values other than "Insurance" or "Self Pay" never reach
  Rosanna or Jasmine.

Optional, per-run overrides (all off by default; nothing below changes the
standard schedule unless the operator turns it on for that run):
- include_programming: bill Programming (Detox, Residential) regardless of
  the weekday, so it can be included on a Monday or Wednesday. E-care is
  unaffected and stays Tuesday-only.
- Cathy report: pull every Professional (CMS-1500/UB-04) Insurance row
  whose Payer is Oxford, ConnectiCare, or UBH (see is_cathy_payer) out of
  the professional pool and assign it to Cathy instead, so a row is never
  worked twice. Whatever the service is, it is hers — including IOP for
  those three payers, which she takes ahead of the IOP-to-Jasmine rule.
  WM, PHP and the O'Flynn Karen rule still take priority over it.
- Cathy report, all of her payers: the same report run against Cathy's
  full payer list (see is_cathy_all_payer / CATHY_ALL_PAYERS) instead of
  just her usual three. Only the payer list widens; it is still
  Professional Insurance rows only, and it turns the Cathy report on by
  itself.
- Nothing for Rosanna: give Rosanna no rows at all for the run. Her share
  of the professional pool goes to Jasmine instead, exactly as it does on
  a weekend, and no workbook is generated for her.
- Exclude Aetna: drop Aetna rows (see is_aetna_payer) from the individual
  staff reports; they stay in the Masters workbook.
"""
import re
from datetime import datetime
from typing import Tuple


DRUG_SCREEN_KEYWORDS = ("drug screen", "utox", "urine tox", "drug test", "uds")
PHP_KEYWORDS = ("partial hospitalization", "php")

# Rosanna's professional-service row cap by weekday (0=Monday..6=Sunday).
# Rosanna works Monday through Friday, capped at 150 rows/day; she receives
# no rows on weekends.
ROSANNA_PROFESSIONAL_CAP = {0: 150, 1: 150, 2: 150, 3: 150, 4: 150}

# Payers that go to Cathy's Professional-services-only report when that
# optional report is turned on. Matching is case-insensitive and tolerant of
# the spelling variants these payers appear with in the Payer column
# ("ConnectiCare"/"Connecti Care", "UBH"/"United Behavioral Health").
CATHY_PAYERS = ("Oxford", "ConnectiCare", "UBH")
_CATHY_PAYER_RE = re.compile(
    r"oxford|connecti[\s\-]?care|\bubh\b|united\s+behavioral\s+health",
    re.IGNORECASE,
)

# The full payer list Cathy can be given for a run: her usual three plus the
# Optum-family plans that sit alongside them in the report's payer filter.
# "UBH-HP" is spelled out here for the operator's benefit, but it already
# matches the UBH pattern above, so it is Cathy's under either payer list.
CATHY_ALL_PAYERS = ("ConnectiCare", "Emblem", "Oxford", "Surest", "UBH",
                    "UBH-HP", "UMR")
_CATHY_ALL_PAYER_RE = re.compile(
    r"oxford|connecti[\s\-]?care|\bubh\b|united\s+behavioral\s+health"
    r"|emblem|surest|\bumr\b",
    re.IGNORECASE,
)


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


def is_php_service(service: str) -> bool:
    """Return True for PHP/Partial Hospitalization services (always Melissa's)."""
    s = (service or "").lower()
    return any(kw in s for kw in PHP_KEYWORDS)


def _is_programming_service(service: str) -> bool:
    """Return True for Programming services: Detox, Residential.

    IOP is intentionally excluded here: it bills every day of the week
    (see is_iop_service) rather than following this weekday schedule.
    PHP is also excluded: it's always assigned to Melissa (see
    is_php_service) rather than following this weekday schedule.
    """
    s = service.lower()
    return (
        "detox" in s
        or "residential" in s
    )


def is_iop_service(service: str) -> bool:
    """Return True for any IOP service, including Telemed IOP (case-insensitive)."""
    return "iop" in (service or "").lower()


def is_professional_claim_type(claim_type: str) -> bool:
    """Return True if the Claim Type column value is CMS-1500 or UB-04 (Professional).

    UB-04 counts as Professional every day of the week, alongside CMS-1500.
    """
    return (claim_type or "").strip().upper() in ("CMS-1500", "UB-04")


def is_cathy_payer(payer: str) -> bool:
    """Return True if the Payer column value is Oxford, ConnectiCare, or UBH.

    Used by the optional Cathy report, which takes only Professional
    (CMS-1500/UB-04) Insurance rows for these three payers.
    """
    return bool(_CATHY_PAYER_RE.search(payer or ""))


def is_cathy_all_payer(payer: str) -> bool:
    """Return True if the Payer column value is on Cathy's full payer list.

    A superset of is_cathy_payer: her usual three payers plus Emblem,
    Surest, UBH-HP and UMR. Used by the optional "all of her payers"
    variant of the Cathy report, which is otherwise identical — still
    Professional (CMS-1500/UB-04) Insurance rows only.
    """
    return bool(_CATHY_ALL_PAYER_RE.search(payer or ""))


def is_aetna_payer(payer: str) -> bool:
    """Return True if the Payer column value refers to Aetna."""
    return "aetna" in (payer or "").lower()


def is_non_billable_service_for_weekday(
    service: str,
    weekday: int,
    is_professional: bool = False,
    self_pay: bool = False,
    include_programming: bool = False,
) -> bool:
    """
    Return True if the given service (string) should be treated as non-billable
    on the specified weekday.

    - service: original service string (case-insensitive checks will be used)
    - weekday: 0=Monday .. 6=Sunday
    - is_professional: True when the row's Claim Type is CMS-1500 or UB-04.
      Professional rows bill every day of the week.
    - self_pay: when True, every service is billable every day, with no
      exceptions (including e-care).
    - include_programming: one-off override. When True, Programming (Detox,
      Residential) is billable regardless of the weekday, so it can be
      included on a Monday or Wednesday. It does not affect e-care, which
      stays Tuesday-only.
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

    # IOP (including Telemed IOP) bills every day of the week.
    if is_iop_service(s):
        return False

    # Programming (Detox, Residential) bills Tue/Thu/Fri + weekends, unless
    # the one-off include_programming override is on for this run.
    if _is_programming_service(s):
        if include_programming:
            return False
        return weekday not in (1, 3, 4, 5, 6)

    # Anything else is neither Professional nor Programming nor e-care, and
    # isn't addressed by the daily schedule, so it is not billable.
    return True
