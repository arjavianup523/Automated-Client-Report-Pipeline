
from math import isclose


def validate_portfolio(holdings, performance_records):
    errors = []

    if not holdings:
        errors.append("No holdings found for the selected reporting period.")

    security_ids = [item.get("security_id") for item in holdings]

    if any(not security_id for security_id in security_ids):
        errors.append("A holding is missing its security ID.")

    if len(security_ids) != len(set(security_ids)):
        errors.append("Duplicate securities found in the portfolio.")

    for item in holdings:
        market_value = item.get("market_value")
        weight = item.get("weight")

        if (
            not isinstance(market_value, (int, float))
            or isinstance(market_value, bool)
            or market_value < 0
        ):
            errors.append(
                f"Invalid market value for {item.get('security_id', 'unknown security')}."
            )

        if (
            not isinstance(weight, (int, float))
            or isinstance(weight, bool)
            or not 0 <= weight <= 1
        ):
            errors.append(
                f"Invalid portfolio weight for {item.get('security_id', 'unknown security')}."
            )

    if holdings and not errors:
        total_weight = sum(item["weight"] for item in holdings)

        if not isclose(total_weight, 1.0, abs_tol=0.000001):
            errors.append(
                f"Portfolio weights total {total_weight:.2%}, not 100%."
            )

    if len(performance_records) != 1:
        errors.append(
            "Expected exactly one performance record for the selected fund and period."
        )

    return {
        "passed": len(errors) == 0,
        "errors": errors,
    }
