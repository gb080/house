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