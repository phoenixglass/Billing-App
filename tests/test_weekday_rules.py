"""
Unit tests for weekday-based non-billable service rules.

Weekly billing schedule:
- Monday:    Professional + Utox     (Programming + e-care non-billable)
- Tuesday:   E-care + Programming    (Professional + Utox non-billable)
- Wednesday: Professional + Utox     (Programming + e-care non-billable)
- Thursday:  Programming + Professional + Utox (e-care non-billable)
- Friday:    Programming + Professional + Utox (e-care non-billable)
- Sat/Sun:   all services billable except e-care

Programming = Detox, Residential, Partial Hospitalization (PHP), IOP
Professional = all other services (including Acupuncture)
Utox = drug screen services
E-care is billable on Tuesdays only.
"""
import sys
from pathlib import Path

# Add parent directory to path to import billing_rules
sys.path.insert(0, str(Path(__file__).parent.parent))

from billing_rules import (
    _is_drug_screen,
    _is_ecare,
    _is_programming_service,
    is_non_billable_service_for_weekday,
)


def test_ecare_variants():
    """Test that all e-care variants are recognized."""
    # Test various spellings (case-insensitive)
    assert _is_ecare('e-care')
    assert _is_ecare('e care')
    assert _is_ecare('ecare')
    assert _is_ecare('E-Care')
    assert _is_ecare('E Care')
    assert _is_ecare('ECare')
    assert _is_ecare('extended care')
    assert _is_ecare('Extended Care')
    assert _is_ecare('EXTENDED CARE')

    # Test in context
    assert _is_ecare('service e-care session')
    assert _is_ecare('telehealth e care')
    assert _is_ecare('telehealth extended care')

    # Test non-e-care services
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

    # Non-drug-screen services
    assert not _is_drug_screen('Detox')
    assert not _is_drug_screen('IOP')
    assert not _is_drug_screen('Acupuncture')


def test_programming_service_classification():
    """Test that Programming services (Detox, Residential, PHP, IOP) are recognized."""
    assert _is_programming_service('Detox')
    assert _is_programming_service('Residential')
    assert _is_programming_service('Partial Hospitalization')
    assert _is_programming_service('PHP')
    assert _is_programming_service('IOP')
    assert _is_programming_service('DETOX SERVICES')

    # Professional services are not Programming
    assert not _is_programming_service('Acupuncture')
    assert not _is_programming_service('Individual Therapy')


def test_monday_professional_and_utox():
    """Monday: Professional + Utox billable; Programming + e-care non-billable."""
    weekday = 0  # Monday

    # Programming services are non-billable on Monday (now includes IOP)
    assert is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert is_non_billable_service_for_weekday('Residential', weekday)
    assert is_non_billable_service_for_weekday('Detox', weekday)
    assert is_non_billable_service_for_weekday('IOP', weekday)

    # Professional services are billable on Monday
    assert not is_non_billable_service_for_weekday('Acupuncture', weekday)
    assert not is_non_billable_service_for_weekday('Individual Therapy', weekday)

    # Utox/drug screens are billable on Monday
    assert not is_non_billable_service_for_weekday('Drug Screen', weekday)
    assert not is_non_billable_service_for_weekday('Utox', weekday)

    # E-care is non-billable on Monday
    assert is_non_billable_service_for_weekday('e-care', weekday)
    assert is_non_billable_service_for_weekday('extended care', weekday)


def test_tuesday_ecare_and_programming():
    """Tuesday: E-care + Programming billable; Professional + Utox non-billable."""
    weekday = 1  # Tuesday

    # Programming services are billable on Tuesday (now includes IOP)
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)
    assert not is_non_billable_service_for_weekday('Detox', weekday)
    assert not is_non_billable_service_for_weekday('IOP', weekday)

    # Professional services are non-billable on Tuesday
    assert is_non_billable_service_for_weekday('Acupuncture', weekday)
    assert is_non_billable_service_for_weekday('Individual Therapy', weekday)

    # Utox/drug screens are non-billable on Tuesday
    assert is_non_billable_service_for_weekday('Drug Screen', weekday)
    assert is_non_billable_service_for_weekday('Utox', weekday)

    # E-care is billable on Tuesday
    assert not is_non_billable_service_for_weekday('e-care', weekday)
    assert not is_non_billable_service_for_weekday('extended care', weekday)


def test_wednesday_professional_and_utox():
    """Wednesday: Professional + Utox billable; Programming + e-care non-billable."""
    weekday = 2  # Wednesday

    # Programming services are non-billable on Wednesday (now includes IOP)
    assert is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert is_non_billable_service_for_weekday('Residential', weekday)
    assert is_non_billable_service_for_weekday('Detox', weekday)
    assert is_non_billable_service_for_weekday('IOP', weekday)

    # Professional services are billable on Wednesday
    assert not is_non_billable_service_for_weekday('Acupuncture', weekday)

    # Utox/drug screens are billable on Wednesday
    assert not is_non_billable_service_for_weekday('Drug Screen', weekday)
    assert not is_non_billable_service_for_weekday('Utox', weekday)

    # E-care is non-billable on Wednesday
    assert is_non_billable_service_for_weekday('e-care', weekday)
    assert is_non_billable_service_for_weekday('extended care', weekday)


def test_thursday_programming_professional_utox():
    """Thursday: Programming + Professional + Utox billable; e-care non-billable."""
    weekday = 3  # Thursday

    # All three categories are billable on Thursday
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)
    assert not is_non_billable_service_for_weekday('Detox', weekday)
    assert not is_non_billable_service_for_weekday('IOP', weekday)
    assert not is_non_billable_service_for_weekday('Acupuncture', weekday)
    assert not is_non_billable_service_for_weekday('Drug Screen', weekday)
    assert not is_non_billable_service_for_weekday('Utox', weekday)

    # E-care is non-billable on Thursday
    assert is_non_billable_service_for_weekday('e-care', weekday)
    assert is_non_billable_service_for_weekday('extended care', weekday)


def test_friday_programming_professional_utox():
    """Friday: Programming + Professional + Utox billable; e-care non-billable."""
    weekday = 4  # Friday

    # All three categories are billable on Friday
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)
    assert not is_non_billable_service_for_weekday('Detox', weekday)
    assert not is_non_billable_service_for_weekday('IOP', weekday)
    assert not is_non_billable_service_for_weekday('Acupuncture', weekday)
    assert not is_non_billable_service_for_weekday('Drug Screen', weekday)
    assert not is_non_billable_service_for_weekday('Utox', weekday)

    # E-care is non-billable on Friday
    assert is_non_billable_service_for_weekday('e-care', weekday)
    assert is_non_billable_service_for_weekday('extended care', weekday)


def test_weekend_only_ecare_non_billable():
    """Saturday and Sunday: all services billable except e-care."""
    for weekday in (5, 6):
        assert is_non_billable_service_for_weekday('e-care', weekday)
        assert is_non_billable_service_for_weekday('extended care', weekday)
        assert not is_non_billable_service_for_weekday('Detox', weekday)
        assert not is_non_billable_service_for_weekday('Residential', weekday)
        assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
        assert not is_non_billable_service_for_weekday('IOP', weekday)
        assert not is_non_billable_service_for_weekday('Acupuncture', weekday)
        assert not is_non_billable_service_for_weekday('Drug Screen', weekday)
        assert not is_non_billable_service_for_weekday('Utox', weekday)


def test_php_on_monday_option():
    """php_on_monday=True makes Partial Hospitalization billable on Mondays."""
    weekday = 0  # Monday

    # Default: PHP is non-billable on Monday
    assert is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert is_non_billable_service_for_weekday('partial hospitalization', weekday)
    assert is_non_billable_service_for_weekday('PHP', weekday)

    # With php_on_monday=True: PHP is billable on Monday
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday, php_on_monday=True)
    assert not is_non_billable_service_for_weekday('partial hospitalization', weekday, php_on_monday=True)
    assert not is_non_billable_service_for_weekday('PHP', weekday, php_on_monday=True)

    # php_on_monday=True does NOT affect other Monday restrictions
    assert is_non_billable_service_for_weekday('Residential', weekday, php_on_monday=True)
    assert is_non_billable_service_for_weekday('Detox', weekday, php_on_monday=True)
    assert is_non_billable_service_for_weekday('IOP', weekday, php_on_monday=True)
    assert is_non_billable_service_for_weekday('e-care', weekday, php_on_monday=True)

    # php_on_monday only applies to Monday; PHP stays non-billable on Wednesday
    assert is_non_billable_service_for_weekday('Partial Hospitalization', 2, php_on_monday=True)
    # PHP is already billable Tue/Thu/Fri regardless of the flag
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', 1, php_on_monday=True)
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', 3, php_on_monday=True)
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', 4, php_on_monday=True)


def test_case_insensitive_matching():
    """Test that service matching is case-insensitive."""
    weekday = 0  # Monday

    # Programming services are non-billable on Monday regardless of case
    assert is_non_billable_service_for_weekday('DETOX', weekday)
    assert is_non_billable_service_for_weekday('detox', weekday)
    assert is_non_billable_service_for_weekday('Detox', weekday)

    assert is_non_billable_service_for_weekday('RESIDENTIAL', weekday)
    assert is_non_billable_service_for_weekday('residential', weekday)

    assert is_non_billable_service_for_weekday('PARTIAL HOSPITALIZATION', weekday)
    assert is_non_billable_service_for_weekday('partial hospitalization', weekday)

    assert is_non_billable_service_for_weekday('iop', weekday)
    assert is_non_billable_service_for_weekday('IOP', weekday)


if __name__ == '__main__':
    # Run tests
    test_ecare_variants()
    print("✓ test_ecare_variants passed")

    test_drug_screen_variants()
    print("✓ test_drug_screen_variants passed")

    test_programming_service_classification()
    print("✓ test_programming_service_classification passed")

    test_monday_professional_and_utox()
    print("✓ test_monday_professional_and_utox passed")

    test_tuesday_ecare_and_programming()
    print("✓ test_tuesday_ecare_and_programming passed")

    test_wednesday_professional_and_utox()
    print("✓ test_wednesday_professional_and_utox passed")

    test_thursday_programming_professional_utox()
    print("✓ test_thursday_programming_professional_utox passed")

    test_friday_programming_professional_utox()
    print("✓ test_friday_programming_professional_utox passed")

    test_weekend_only_ecare_non_billable()
    print("✓ test_weekend_only_ecare_non_billable passed")

    test_php_on_monday_option()
    print("✓ test_php_on_monday_option passed")

    test_case_insensitive_matching()
    print("✓ test_case_insensitive_matching passed")

    print("\nAll tests passed!")
