
from src.validation import validate_portfolio


def valid_holdings():
    return [
        {"security_id": "SEC001", "market_value": 3000000, "weight": 0.30},
        {"security_id": "SEC002", "market_value": 2500000, "weight": 0.25},
        {"security_id": "SEC003", "market_value": 2000000, "weight": 0.20},
        {"security_id": "SEC004", "market_value": 1500000, "weight": 0.15},
        {"security_id": "SEC005", "market_value": 1000000, "weight": 0.10},
    ]


def test_valid_portfolio_passes():
    result = validate_portfolio(valid_holdings(), [{"fund_return": 0.042}])
    assert result["passed"] is True
    assert result["errors"] == []


def test_incorrect_weights_fail():
    holdings = valid_holdings()
    holdings[0]["weight"] = 0.20

    result = validate_portfolio(holdings, [{"fund_return": 0.042}])

    assert result["passed"] is False
    assert any("not 100%" in error for error in result["errors"])


def test_negative_market_value_fails():
    holdings = valid_holdings()
    holdings[0]["market_value"] = -100

    result = validate_portfolio(holdings, [{"fund_return": 0.042}])

    assert result["passed"] is False
    assert any("Invalid market value" in error for error in result["errors"])


def test_duplicate_security_fails():
    holdings = valid_holdings()
    holdings[1]["security_id"] = holdings[0]["security_id"]

    result = validate_portfolio(holdings, [{"fund_return": 0.042}])

    assert result["passed"] is False
    assert any("Duplicate securities" in error for error in result["errors"])


def test_missing_performance_record_fails():
    result = validate_portfolio(valid_holdings(), [])

    assert result["passed"] is False
    assert any(
        "exactly one performance record" in error
        for error in result["errors"]
    )
