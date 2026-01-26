"""
Unit tests for weekday-based non-billable service rules.
"""
import sys
from pathlib import Path

# Add parent directory to path to import billing_rules
sys.path.insert(0, str(Path(__file__).parent.parent))

from billing_rules import _is_ecare, is_non_billable_service_for_weekday


def test_ecare_variants():
    """Test that all e-care variants are recognized."""
    # Test various spellings (case-insensitive)
    assert _is_ecare('e-care')
    assert _is_ecare('e care')
    assert _is_ecare('ecare')
    assert _is_ecare('E-Care')
    assert _is_ecare('E Care')
    assert _is_ecare('ECare')
    
    # Test in context
    assert _is_ecare('service e-care session')
    assert _is_ecare('telehealth e care')
    
    # Test non-e-care services
    assert not _is_ecare('detox')
    assert not _is_ecare('residential')
    assert not _is_ecare('iop')


def test_tuesday_bills_everything():
    """Test that Tuesday (weekday=1) bills everything including e-care."""
    weekday = 1  # Tuesday
    
    # E-care should be billable (not non-billable) on Tuesday
    assert not is_non_billable_service_for_weekday('e-care', weekday)
    assert not is_non_billable_service_for_weekday('e care', weekday)
    assert not is_non_billable_service_for_weekday('ecare', weekday)
    
    # All other services should also be billable
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)
    assert not is_non_billable_service_for_weekday('Detox', weekday)
    assert not is_non_billable_service_for_weekday('IOP', weekday)


def test_monday_non_billable():
    """Test that Monday (weekday=0) marks specific services as non-billable."""
    weekday = 0  # Monday
    
    # E-care should be non-billable on Monday
    assert is_non_billable_service_for_weekday('e-care', weekday)
    assert is_non_billable_service_for_weekday('E Care', weekday)
    assert is_non_billable_service_for_weekday('ecare session', weekday)
    
    # These services should be non-billable on Monday
    assert is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert is_non_billable_service_for_weekday('Residential', weekday)
    assert is_non_billable_service_for_weekday('Detox', weekday)
    
    # Other services should be billable on Monday
    assert not is_non_billable_service_for_weekday('IOP', weekday)
    assert not is_non_billable_service_for_weekday('Acupuncture', weekday)


def test_wednesday_only_ecare_non_billable():
    """Test that Wednesday (weekday=2) only marks e-care as non-billable."""
    weekday = 2  # Wednesday
    
    # E-care should be non-billable
    assert is_non_billable_service_for_weekday('e-care', weekday)
    assert is_non_billable_service_for_weekday('E Care', weekday)
    assert is_non_billable_service_for_weekday('ecare', weekday)
    
    # All other services should be billable on Wednesday
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)
    assert not is_non_billable_service_for_weekday('Detox', weekday)
    assert not is_non_billable_service_for_weekday('IOP', weekday)


def test_thursday_only_ecare_non_billable():
    """Test that Thursday (weekday=3) only marks e-care as non-billable."""
    weekday = 3  # Thursday
    
    # E-care should be non-billable
    assert is_non_billable_service_for_weekday('e-care', weekday)
    
    # All other services should be billable
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)
    assert not is_non_billable_service_for_weekday('Detox', weekday)


def test_friday_only_ecare_non_billable():
    """Test that Friday (weekday=4) only marks e-care as non-billable."""
    weekday = 4  # Friday
    
    # E-care should be non-billable
    assert is_non_billable_service_for_weekday('e-care', weekday)
    
    # All other services should be billable
    assert not is_non_billable_service_for_weekday('Partial Hospitalization', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)


def test_weekend_only_ecare_non_billable():
    """Test that Saturday and Sunday only mark e-care as non-billable."""
    # Saturday
    weekday = 5
    assert is_non_billable_service_for_weekday('e-care', weekday)
    assert not is_non_billable_service_for_weekday('Detox', weekday)
    
    # Sunday
    weekday = 6
    assert is_non_billable_service_for_weekday('e-care', weekday)
    assert not is_non_billable_service_for_weekday('Residential', weekday)


def test_case_insensitive_matching():
    """Test that service matching is case-insensitive."""
    weekday = 0  # Monday
    
    # Test various case combinations
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
    
    test_tuesday_bills_everything()
    print("✓ test_tuesday_bills_everything passed")
    
    test_monday_non_billable()
    print("✓ test_monday_non_billable passed")
    
    test_wednesday_only_ecare_non_billable()
    print("✓ test_wednesday_only_ecare_non_billable passed")
    
    test_thursday_only_ecare_non_billable()
    print("✓ test_thursday_only_ecare_non_billable passed")
    
    test_friday_only_ecare_non_billable()
    print("✓ test_friday_only_ecare_non_billable passed")
    
    test_weekend_only_ecare_non_billable()
    print("✓ test_weekend_only_ecare_non_billable passed")
    
    test_case_insensitive_matching()
    print("✓ test_case_insensitive_matching passed")
    
    print("\nAll tests passed!")
