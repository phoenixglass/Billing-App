"""
Billing rules for weekday-based non-billable service logic.
"""
from datetime import datetime


def _is_ecare(service: str) -> bool:
    """
    Check if a service is e-care (case-insensitive).
    
    Args:
        service: Service name (will be converted to lowercase)
        
    Returns:
        True if the service is e-care in any variant
    """
    # Check for common e-care variants (case-insensitive)
    service_lower = service.lower()
    ecare_variants = ['e-care', 'e care', 'ecare']
    return any(variant in service_lower for variant in ecare_variants)


def is_non_billable_service_for_weekday(service: str, weekday: int) -> bool:
    """
    Determine if a service is non-billable based on the weekday.
    
    Weekday rules:
    - Tuesday (weekday == 1): Everything is billed (including e-care)
    - Monday (weekday == 0): Non-billable: partial hospitalization, residential, detox, and e-care
    - Wednesday-Friday (2, 3, 4): All services billed except e-care
    - Saturday-Sunday (5, 6): All services billed except e-care
    
    E-care is only billable on Tuesdays (non-billable all other days).
    
    Args:
        service: The service name (case-insensitive matching)
        weekday: Day of week (0=Monday, 1=Tuesday, ..., 6=Sunday)
        
    Returns:
        True if the service is non-billable for the given weekday, False otherwise
    """
    service_lower = service.lower()
    
    # Tuesday (1): Bill everything - nothing is non-billable
    if weekday == 1:
        return False
    
    # Monday (0): Non-billable for partial hospitalization, residential, detox, and e-care
    if weekday == 0:
        if _is_ecare(service_lower):
            return True
        if 'partial hospitalization' in service_lower:
            return True
        if 'residential' in service_lower:
            return True
        if 'detox' in service_lower:
            return True
        return False
    
    # Wednesday-Sunday (2, 3, 4, 5, 6): Non-billable only for e-care
    if weekday in [2, 3, 4, 5, 6]:
        return _is_ecare(service_lower)
    
    return False
