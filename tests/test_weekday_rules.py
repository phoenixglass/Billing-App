import billing_rules as br

def test_ecare_variants_monday():
    # Monday (0) — e-care should be non-billable
    assert br.is_non_billable_service_for_weekday("E-Care Individual", 0) is True
    assert br.is_non_billable_service_for_weekday("e care group", 0) is True
    assert br.is_non_billable_service_for_weekday("ecare session", 0) is True

def test_ecare_variants_tuesday():
    # Tuesday (1) — e-care should be billable (so non-billable is False)
    assert br.is_non_billable_service_for_weekday("E-Care Individual", 1) is False
    assert br.is_non_billable_service_for_weekday("e care group", 1) is False

def test_monday_exclusions():
    # Monday excludes partial hospitalization/residential/detox
    assert br.is_non_billable_service_for_weekday("Partial Hospitalization - day", 0) is True
    assert br.is_non_billable_service_for_weekday("Residential Program", 0) is True
    assert br.is_non_billable_service_for_weekday("Detox", 0) is True

def test_wed_through_fri_only_ecare_nonbillable():
    # Wednesday (2) — only e-care should be non-billable
    assert br.is_non_billable_service_for_weekday("Regular Therapy", 2) is False
    assert br.is_non_billable_service_for_weekday("E-Care", 2) is True
