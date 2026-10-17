FULL_DAY_RATES = {
    "Labor": 500.0,
    "HED LABOR": 250.0,
    "Skill": 650.0,
    "Forman": 800.0,
    "Mason": 800.0,
    "Carpenter": 900.0,
    "Electrician": 1000.0,
    "Plumber": 1000.0,
    "Painter": 800.0,
}

def calculate_labor_pay(worked_days, role, daily_rate=None):
    days = max(0.0, float(worked_days))
    rate = FULL_DAY_RATES.get(role, 0.0) if daily_rate is None else max(0.0, float(daily_rate))
    full_days = int(days)
    partial_days = days - full_days
    full_pay = full_days * rate
    partial_pay = partial_days * rate
    return full_pay + partial_pay, full_pay, partial_pay