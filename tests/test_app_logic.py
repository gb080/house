from app_logic import calculate_labor_pay


def test_calculate_labor_pay_splits_full_and_partial_days():
    assert calculate_labor_pay(1.5, "Mason") == (1200.0, 800.0, 400.0)


def test_calculate_labor_pay_clamps_negative_days():
    assert calculate_labor_pay(-2, "Mason") == (0.0, 0.0, 0.0)


def test_calculate_labor_pay_uses_zero_for_unknown_role():
    assert calculate_labor_pay(2, "Unknown") == (0.0, 0.0, 0.0)


def test_displayed_skill_and_forman_roles_have_rates():
    assert calculate_labor_pay(1, "Skill")[0] == 650.0
    assert calculate_labor_pay(1, "Forman")[0] == 800.0


def test_hed_labor_role_rate_is_250_per_day():
    assert calculate_labor_pay(1, "HED LABOR")[0] == 250.0


def test_labor_role_keeps_its_existing_rate():
    assert calculate_labor_pay(1, "Labor")[0] == 500.0


def test_custom_labor_rate_calculates_full_and_partial_days():
    assert calculate_labor_pay(1.5, "Custom role", 300.0) == (450.0, 300.0, 150.0)