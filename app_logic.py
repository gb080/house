FULL_DAY_RATES = {
    "Labor": 600.0,
    "Mason": 800.0,
    "Carpenter": 900.0,
    "Electrician": 1000.0,
    "Plumber": 1000.0,
    "Painter": 800.0,
}

TIER_TABLE = {
    "full_day": 1.0,
    "half_day": 0.5,
    "quarter_day": 0.25,
}


def get_partial_rate(days, role):
    return FULL_DAY_RATES.get(role, 0.0) * max(0.0, float(days))


def calculate_labor_pay(worked_days, role):
    days = max(0.0, float(worked_days))
    rate = FULL_DAY_RATES.get(role, 0.0)
    full_days = int(days)
    partial_days = days - full_days
    full_pay = full_days * rate
    partial_pay = partial_days * rate
    return full_pay + partial_pay, full_pay, partial_pay