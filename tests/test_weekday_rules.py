"""
Unit tests for the daily billing schedule helpers in billing_rules.py.

Daily billing schedule:
- Professional (Claim Type CMS-1500) services bill every day of the week.
- IOP (including Telemed IOP) bills every day of the week, with no
  exceptions.
- Programming (Detox, Residential) bills Tuesday, Thursday, Friday, and
  weekends; non-billable Monday and Wednesday.
- E-care bills on Tuesdays only, regardless of Claim Type.
- Self Pay bills every service every day, with no exceptions.
- PHP (Partial Hospitalization) always goes to Melissa and isn't governed
  by this weekday schedule at all.

Rosanna/Joshua/Jasmine split (Insurance rows only):
- Rosanna's professional-service row cap by weekday: Monday=300 only.
- Joshua's professional-service row cap by weekday: Tuesday=300,
  Thursday=125, Friday=125.
- Wednesday and weekends have no cap for either of them (Jasmine gets the
  whole professional pool those days).
"""
import sys
from pathlib import Path

# Add parent directory to path to import billing_rules
sys.path.insert(0, str(Path(__file__).parent.parent))

from billing_rules import (
    _is_drug_screen,
    _is_ecare,
    _is_programming_service,
    is_iop_service,
    is_non_billable_service_for_weekday,
    is_php_service,
    is_professional_claim_type,
    parse_weekday_from_token,
    JOSHUA_PROFESSIONAL_CAP,
    ROSANNA_PROFESSIONAL_CAP,
)


def test_ecare_variants():
    """Test that all e-care variants are recognized."""
    assert _is_ecare('e-care')
    assert _is_ecare('e care')
    assert _is_ecare('ecare')
    assert _is_ecare('E-Care')
    assert _is_ecare('E Care')
    assert _is_ecare('ECare')
    assert _is_ecare('extended care')
    assert _is_ecare('Extended Care')
    assert _is_ecare('EXTENDED CARE')

    assert _is_ecare('service e-care session')
    assert _is_ecare('telehealth e care')
    assert _is_ecare('telehealth extended care')

    assert not _is_ecare('detox')
    assert not _is_ecare('residential')
    assert not _is_ecare('iop')


def test_drug_screen_variants():
    """Test that Utox / drug screen variants are recognized."""
    assert _is_drug_screen('Drug Screen')
    assert _is_drug_screen('drug screen')
    assert _is_drug_screen('Utox')
    assert _is_drug_screen('UTOX')
    assert _is_drug_screen('Urine Tox')
    assert _is_drug_screen('Drug Test')
    assert _is_drug_screen('UDS')

    assert not _is_drug_screen('Detox')
    assert not _is_drug_screen('IOP')
    assert not _is_drug_screen('Acupuncture')


def test_programming_service_classification():
    """Test that Programming services (Detox, Residential) are recognized.

    IOP is intentionally excluded: it bills every day of the week (see
    is_iop_service) rather than following the Programming schedule. PHP is
    also excluded: it always goes to Melissa (see is_php_service) rather
    than following the Programming schedule.
    """
    assert _is_programming_service('Detox')
    assert _is_programming_service('Residential')
    assert _is_programming_service('DETOX SERVICES')

    assert not _is_programming_service('IOP')
    assert not _is_programming_service('Partial Hospitalization')
    assert not _is_programming_service('PHP')
    assert not _is_programming_service('Acupuncture')
    assert not _is_programming_service('Individual Therapy')


def test_is_php_service_variants():
    """Test that PHP / Partial Hospitalization is recognized regardless of case."""
    assert is_php_service('PHP')
    assert is_php_service('php')
    assert is_php_service('Partial Hospitalization')
    assert is_php_service('PARTIAL HOSPITALIZATION')
    assert is_php_service('Partial Hospitalization - Program')

    assert not is_php_service('Detox')
    assert not is_php_service('Residential')
    assert not is_php_service('IOP')
    assert not is_php_service('')
    assert not is_php_service(None)


def test_is_iop_service_variants():
    """Test that IOP, including Telemed IOP, is recognized regardless of case."""
    assert is_iop_service('IOP')
    assert is_iop_service('iop')
    assert is_iop_service('Telemed IOP')
    assert is_iop_service('TELEMED IOP')
    assert is_iop_service('IOP - Group')

    assert not is_iop_service('Detox')
    assert not is_iop_service('Residential')
    assert not is_iop_service('')
    assert not is_iop_service(None)


def test_claim_type_professional():
    """Only Claim Type == CMS-1500 (case/whitespace-insensitive) is Professional."""
    assert is_professional_claim_type('CMS-1500')
    assert is_professional_claim_type('cms-1500')
    assert is_professional_claim_type('  CMS-1500  ')

    assert not is_professional_claim_type('UB-04')
    assert not is_professional_claim_type('')
    assert not is_professional_claim_type(None)


def test_professional_bills_every_day():
    """A row flagged is_professional=True is billable on every weekday."""
    for weekday in range(7):
        assert not is_non_billable_service_for_weekday('Individual Therapy', weekday, is_professional=True)
        assert not is_non_billable_service_for_weekday('Detox', weekday, is_professional=True)
        # Even e-care bills every day once it's flagged Professional.
        assert not is_non_billable_service_for_weekday('e-care', weekday, is_professional=True)


def test_monday_and_wednesday_programming_non_billable():
    """Monday and Wednesday: Programming and e-care are non-billable, but IOP is billable."""
    for weekday in (0, 2):
        assert is_non_billable_service_for_weekday('Detox', weekday)
        assert is_non_billable_service_for_weekday('Residential', weekday)
        assert not is_non_billable_service_for_weekday('IOP', weekday)
        assert is_non_billable_service_for_weekday('e-care', weekday)
        assert is_non_billable_service_for_weekday('extended care', weekday)

        # Non-professional, non-programming services aren't addressed by the
        # schedule at all, so they're non-billable too.
        assert is_non_billable_service_for_weekday('Acupuncture', weekday)


def test_tuesday_programming_and_ecare_billable():
    """Tuesday: Programming (including e-care) is billable."""
    weekday = 1
    assert not is_non_billable_service_for_weekday('Detox', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)
    assert not is_non_billable_service_for_weekday('IOP', weekday)
    assert not is_non_billable_service_for_weekday('e-care', weekday)
    assert not is_non_billable_service_for_weekday('extended care', weekday)


def test_thursday_friday_programming_billable_no_ecare():
    """Thursday and Friday: Programming billable, e-care is not."""
    for weekday in (3, 4):
        assert not is_non_billable_service_for_weekday('Detox', weekday)
        assert not is_non_billable_service_for_weekday('Residential', weekday)
        assert not is_non_billable_service_for_weekday('IOP', weekday)
        assert is_non_billable_service_for_weekday('e-care', weekday)
        assert is_non_billable_service_for_weekday('extended care', weekday)


def test_weekend_programming_billable_no_ecare():
    """Saturday and Sunday: Programming billable, e-care is not."""
    for weekday in (5, 6):
        assert not is_non_billable_service_for_weekday('Detox', weekday)
        assert not is_non_billable_service_for_weekday('Residential', weekday)
        assert not is_non_billable_service_for_weekday('IOP', weekday)
        assert is_non_billable_service_for_weekday('e-care', weekday)
        assert is_non_billable_service_for_weekday('extended care', weekday)


def test_iop_billable_every_day():
    """IOP (including Telemed IOP) is billable every day of the week, no exceptions."""
    for weekday in range(7):
        assert not is_non_billable_service_for_weekday('IOP', weekday)
        assert not is_non_billable_service_for_weekday('Telemed IOP', weekday)
        assert not is_non_billable_service_for_weekday('IOP - Group', weekday)


def test_self_pay_every_service_every_day_no_exceptions():
    """self_pay=True: every service billable every day, no exceptions (including e-care)."""
    for weekday in range(7):
        assert not is_non_billable_service_for_weekday('Detox', weekday, self_pay=True)
        assert not is_non_billable_service_for_weekday('Residential', weekday, self_pay=True)
        assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday, self_pay=True)
        assert not is_non_billable_service_for_weekday('IOP', weekday, self_pay=True)
        assert not is_non_billable_service_for_weekday('Acupuncture', weekday, self_pay=True)
        assert not is_non_billable_service_for_weekday('Individual Therapy', weekday, self_pay=True)
        assert not is_non_billable_service_for_weekday('Drug Screen', weekday, self_pay=True)
        assert not is_non_billable_service_for_weekday('Utox', weekday, self_pay=True)
        assert not is_non_billable_service_for_weekday('e-care', weekday, self_pay=True)
        assert not is_non_billable_service_for_weekday('extended care', weekday, self_pay=True)


def test_rosanna_professional_cap_by_weekday():
    """Rosanna's professional cap: Monday=300 only."""
    assert ROSANNA_PROFESSIONAL_CAP[0] == 300  # Monday

    # Every other day is intentionally absent -> cap of 0.
    assert ROSANNA_PROFESSIONAL_CAP.get(1, 0) == 0
    assert ROSANNA_PROFESSIONAL_CAP.get(2, 0) == 0
    assert ROSANNA_PROFESSIONAL_CAP.get(3, 0) == 0
    assert ROSANNA_PROFESSIONAL_CAP.get(4, 0) == 0
    assert ROSANNA_PROFESSIONAL_CAP.get(5, 0) == 0
    assert ROSANNA_PROFESSIONAL_CAP.get(6, 0) == 0


def test_joshua_professional_cap_by_weekday():
    """Joshua's professional cap: Tue=300, Thu/Fri=125, Mon/Wed/weekend=0 (absent)."""
    assert JOSHUA_PROFESSIONAL_CAP[1] == 300  # Tuesday
    assert JOSHUA_PROFESSIONAL_CAP[3] == 125  # Thursday
    assert JOSHUA_PROFESSIONAL_CAP[4] == 125  # Friday

    # Monday, Wednesday, and weekends are intentionally absent -> cap of 0.
    assert JOSHUA_PROFESSIONAL_CAP.get(0, 0) == 0
    assert JOSHUA_PROFESSIONAL_CAP.get(2, 0) == 0
    assert JOSHUA_PROFESSIONAL_CAP.get(5, 0) == 0
    assert JOSHUA_PROFESSIONAL_CAP.get(6, 0) == 0


def test_parse_weekday_from_token():
    """MMDDYYYY tokens parse to the correct weekday; invalid tokens fall back."""
    weekday, did_fallback = parse_weekday_from_token('07062026')  # Monday
    assert weekday == 0
    assert not did_fallback

    weekday, did_fallback = parse_weekday_from_token('07072026')  # Tuesday
    assert weekday == 1
    assert not did_fallback

    _, did_fallback = parse_weekday_from_token('not-a-date')
    assert did_fallback

    _, did_fallback = parse_weekday_from_token(None)
    assert did_fallback


def test_case_insensitive_matching():
    """Test that service matching is case-insensitive."""
    weekday = 0  # Monday

    assert is_non_billable_service_for_weekday('DETOX', weekday)
    assert is_non_billable_service_for_weekday('detox', weekday)
    assert is_non_billable_service_for_weekday('Detox', weekday)

    assert is_non_billable_service_for_weekday('RESIDENTIAL', weekday)
    assert is_non_billable_service_for_weekday('residential', weekday)

    assert is_non_billable_service_for_weekday('PARTIAL HOSPITALIZATION', weekday)
    assert is_non_billable_service_for_weekday('partial hospitalization', weekday)

    assert not is_non_billable_service_for_weekday('iop', weekday)
    assert not is_non_billable_service_for_weekday('IOP', weekday)


if __name__ == '__main__':
    test_ecare_variants()
    print("✓ test_ecare_variants passed")

    test_drug_screen_variants()
    print("✓ test_drug_screen_variants passed")

    test_programming_service_classification()
    print("✓ test_programming_service_classification passed")

    test_is_php_service_variants()
    print("✓ test_is_php_service_variants passed")

    test_is_iop_service_variants()
    print("✓ test_is_iop_service_variants passed")

    test_claim_type_professional()
    print("✓ test_claim_type_professional passed")

    test_professional_bills_every_day()
    print("✓ test_professional_bills_every_day passed")

    test_monday_and_wednesday_programming_non_billable()
    print("✓ test_monday_and_wednesday_programming_non_billable passed")

    test_tuesday_programming_and_ecare_billable()
    print("✓ test_tuesday_programming_and_ecare_billable passed")

    test_thursday_friday_programming_billable_no_ecare()
    print("✓ test_thursday_friday_programming_billable_no_ecare passed")

    test_weekend_programming_billable_no_ecare()
    print("✓ test_weekend_programming_billable_no_ecare passed")

    test_iop_billable_every_day()
    print("✓ test_iop_billable_every_day passed")

    test_self_pay_every_service_every_day_no_exceptions()
    print("✓ test_self_pay_every_service_every_day_no_exceptions passed")

    test_rosanna_professional_cap_by_weekday()
    print("✓ test_rosanna_professional_cap_by_weekday passed")

    test_joshua_professional_cap_by_weekday()
    print("✓ test_joshua_professional_cap_by_weekday passed")

    test_parse_weekday_from_token()
    print("✓ test_parse_weekday_from_token passed")

    test_case_insensitive_matching()
    print("✓ test_case_insensitive_matching passed")

    print("\nAll tests passed!")
