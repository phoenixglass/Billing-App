"""
Unit tests for weekday-based non-billable service rules.

Weekly billing schedule:
- Monday:    Professional only        (Programming non-billable)
- Tuesday:   Programming only         (Professional non-billable)
- Wednesday: Professional only        (Programming non-billable)
- Thursday:  Programming + Professional
- Friday:    Programming + Professional
- Sat/Sun:   all services billable except e-care

Programming = Detox, Residential, Partial Hospitalization (PHP)
Professional = all other services (including IOP and Acupuncture)
E-care is billable on Tuesdays only.
"""
import sys
from pathlib import Path

# Add parent directory to path to import billing_rules
sys.path.insert(0, str(Path(__file__).parent.parent))

from billing_rules import (
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


def test_programming_service_classification():
    """Test that Programming services (Detox, Residential, PHP) are recognized."""
    assert _is_programming_service('Detox')
    assert _is_programming_service('Residential')
    assert _is_programming_service('Partial Hospitalization')
    assert _is_programming_service('PHP')
    assert _is_programming_service('DETOX SERVICES')

    # Professional services are not Programming
    assert not _is_programming_service('IOP')
    assert not _is_programming_service('Acupuncture')
    assert not _is_programming_service('Individual Therapy')


def test_monday_professional_only():
    """Monday: Professional billable, Programming non-billable, e-care non-billable."""
    weekday = 0  # Monday

    # Programming services are non-billable on Monday
    assert is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert is_non_billable_service_for_weekday('Residential', weekday)
    assert is_non_billable_service_for_weekday('Detox', weekday)

    # Professional services are billable on Monday
    assert not is_non_billable_service_for_weekday('IOP', weekday)
    assert not is_non_billable_service_for_weekday('Acupuncture', weekday)
    assert not is_non_billable_service_for_weekday('Individual Therapy', weekday)

    # E-care is non-billable on Monday
    assert is_non_billable_service_for_weekday('e-care', weekday)
    assert is_non_billable_service_for_weekday('extended care', weekday)


def test_tuesday_programming_only():
    """Tuesday: Programming billable, Professional non-billable, e-care billable."""
    weekday = 1  # Tuesday

    # Programming services are billable on Tuesday
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)
    assert not is_non_billable_service_for_weekday('Detox', weekday)

    # Professional services are non-billable on Tuesday
    assert is_non_billable_service_for_weekday('IOP', weekday)
    assert is_non_billable_service_for_weekday('Acupuncture', weekday)
    assert is_non_billable_service_for_weekday('Individual Therapy', weekday)

    # E-care is billable on Tuesday
    assert not is_non_billable_service_for_weekday('e-care', weekday)
    assert not is_non_billable_service_for_weekday('extended care', weekday)


def test_wednesday_professional_only():
    """Wednesday: Professional billable, Programming non-billable, e-care non-billable."""
    weekday = 2  # Wednesday

    # Programming services are non-billable on Wednesday
    assert is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert is_non_billable_service_for_weekday('Residential', weekday)
    assert is_non_billable_service_for_weekday('Detox', weekday)

    # Professional services are billable on Wednesday
    assert not is_non_billable_service_for_weekday('IOP', weekday)
    assert not is_non_billable_service_for_weekday('Acupuncture', weekday)

    # E-care is non-billable on Wednesday
    assert is_non_billable_service_for_weekday('e-care', weekday)
    assert is_non_billable_service_for_weekday('extended care', weekday)


def test_thursday_programming_and_professional():
    """Thursday: Programming + Professional both billable, e-care non-billable."""
    weekday = 3  # Thursday

    # Both categories are billable on Thursday
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)
    assert not is_non_billable_service_for_weekday('Detox', weekday)
    assert not is_non_billable_service_for_weekday('IOP', weekday)
    assert not is_non_billable_service_for_weekday('Acupuncture', weekday)

    # E-care is non-billable on Thursday
    assert is_non_billable_service_for_weekday('e-care', weekday)
    assert is_non_billable_service_for_weekday('extended care', weekday)


def test_friday_programming_and_professional():
    """Friday: Programming + Professional both billable, e-care non-billable."""
    weekday = 4  # Friday

    # Both categories are billable on Friday
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)
    assert not is_non_billable_service_for_weekday('Detox', weekday)
    assert not is_non_billable_service_for_weekday('IOP', weekday)
    assert not is_non_billable_service_for_weekday('Acupuncture', weekday)

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


if __name__ == '__main__':
    # Run tests
    test_ecare_variants()
    print("✓ test_ecare_variants passed")

    test_programming_service_classification()
    print("✓ test_programming_service_classification passed")

    test_monday_professional_only()
    print("✓ test_monday_professional_only passed")

    test_tuesday_programming_only()
    print("✓ test_tuesday_programming_only passed")

    test_wednesday_professional_only()
    print("✓ test_wednesday_professional_only passed")

    test_thursday_programming_and_professional()
    print("✓ test_thursday_programming_and_professional passed")

    test_friday_programming_and_professional()
    print("✓ test_friday_programming_and_professional passed")

    test_weekend_only_ecare_non_billable()
    print("✓ test_weekend_only_ecare_non_billable passed")

    test_php_on_monday_option()
    print("✓ test_php_on_monday_option passed")

    test_case_insensitive_matching()
    print("✓ test_case_insensitive_matching passed")

    print("\nAll tests passed!")
