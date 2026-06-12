from datetime import date
from decimal import Decimal

import pytest

from src.tax_model import (
    add_months,
    annual_tax,
    event_tax_batches,
    incremental_event_tax,
    installment_status,
    option_income,
    restricted_stock_income,
)


def test_option_income_for_50000_options():
    income, warnings = option_income("17.80", "14.56", 50000)

    assert income == Decimal("162000.00")
    assert warnings == []


def test_restricted_stock_income_uses_employee_level_inputs():
    income, warnings = restricted_stock_income(
        registration_price="13.06",
        unlock_price="18.00",
        unlock_quantity=20000,
        total_granted=50000,
        total_paid="434000.00",
    )

    assert income == Decimal("137000.00")
    assert warnings == []


def test_restricted_stock_income_rejects_company_total_as_invalid_denominator():
    with pytest.raises(ValueError, match="本批解禁数量不得超过员工实际登记总股数"):
        restricted_stock_income(
            registration_price="13.06",
            unlock_price="18.00",
            unlock_quantity=50001,
            total_granted=50000,
            total_paid="434000.00",
        )


def test_annual_tax_uses_decimal_and_rounds_to_cents():
    result = annual_tax("162000.03")

    assert result == {
        "taxable_income": Decimal("162000.03"),
        "rate": Decimal("0.20"),
        "quick_deduction": Decimal("16920.00"),
        "tax": Decimal("15480.01"),
    }


def test_annual_tax_rounds_exact_half_cent_up():
    assert annual_tax("0.50")["tax"] == Decimal("0.02")


@pytest.mark.parametrize(
    ("income", "rate", "quick_deduction", "tax"),
    [
        ("36000", "0.03", "0.00", "1080.00"),
        ("36000.01", "0.10", "2520.00", "1080.00"),
        ("144000", "0.10", "2520.00", "11880.00"),
        ("144000.01", "0.20", "16920.00", "11880.00"),
        ("300000", "0.20", "16920.00", "43080.00"),
        ("300000.01", "0.25", "31920.00", "43080.00"),
        ("420000", "0.25", "31920.00", "73080.00"),
        ("420000.01", "0.30", "52920.00", "73080.00"),
        ("660000", "0.30", "52920.00", "145080.00"),
        ("660000.01", "0.35", "85920.00", "145080.00"),
        ("960000", "0.35", "85920.00", "250080.00"),
        ("960000.01", "0.45", "181920.00", "250080.00"),
    ],
)
def test_annual_tax_bracket_boundaries(income, rate, quick_deduction, tax):
    result = annual_tax(income)

    assert result["rate"] == Decimal(rate)
    assert result["quick_deduction"] == Decimal(quick_deduction)
    assert result["tax"] == Decimal(tax)


def test_annual_tax_rejects_amount_too_large_to_round():
    with pytest.raises(ValueError, match="应纳税所得额超出可计算范围"):
        annual_tax("1e999999")


def test_incremental_event_tax_uses_one_cumulative_quick_deduction():
    first = incremental_event_tax("100000", "0")
    second = incremental_event_tax("162000", first["cumulative_tax"])

    assert first["cumulative_tax"] == Decimal("7480.00")
    assert first["incremental_tax"] == Decimal("7480.00")
    assert second["cumulative_tax"] == Decimal("15480.00")
    assert second["previous_confirmed_tax"] == Decimal("7480.00")
    assert second["incremental_tax"] == Decimal("8000.00")


def test_event_tax_batches_merge_same_year_but_separate_different_years():
    events = [
        {
            "event_id": "E2",
            "employee_id": "A001",
            "event_date": date(2025, 6, 30),
            "income": "62000",
        },
        {
            "event_id": "E1",
            "employee_id": "A001",
            "event_date": date(2025, 1, 31),
            "income": "100000",
        },
        {
            "event_id": "E3",
            "employee_id": "A001",
            "event_date": date(2026, 2, 28),
            "income": "62000",
        },
    ]

    batches = event_tax_batches(events)

    assert [item["event_id"] for item in batches] == ["E1", "E2", "E3"]
    assert batches[0]["cumulative_income"] == Decimal("100000.00")
    assert batches[0]["incremental_tax"] == Decimal("7480.00")
    assert batches[0]["deadline"] == date(2028, 1, 31)
    assert batches[1]["cumulative_income"] == Decimal("162000.00")
    assert batches[1]["previous_confirmed_tax"] == Decimal("7480.00")
    assert batches[1]["incremental_tax"] == Decimal("8000.00")
    assert batches[1]["deadline"] == date(2028, 6, 30)
    assert batches[2]["tax_year"] == 2026
    assert batches[2]["cumulative_income"] == Decimal("62000.00")
    assert batches[2]["incremental_tax"] == Decimal("3680.00")
    assert batches[2]["deadline"] == date(2029, 2, 28)


@pytest.mark.parametrize(
    ("event_id", "employee_id", "message"),
    [
        (None, "A001", "事件编号不能为空"),
        ("   ", "A001", "事件编号不能为空"),
        ("E1", None, "员工编号不能为空"),
        ("E1", "   ", "员工编号不能为空"),
    ],
)
def test_event_tax_batches_rejects_empty_identifiers(
    event_id, employee_id, message
):
    with pytest.raises(ValueError, match=message):
        event_tax_batches(
            [
                {
                    "event_id": event_id,
                    "employee_id": employee_id,
                    "event_date": date(2026, 1, 1),
                    "income": "100",
                }
            ]
        )


@pytest.mark.parametrize(
    ("function", "args", "message"),
    [
        (option_income, ("17.80", "14.56", 0), "行权数量必须大于0"),
        (option_income, ("-1", "14.56", 1), "行权日市场价不得为负数"),
        (
            restricted_stock_income,
            ("13.06", "18", 1, 0, "0"),
            "员工实际登记总股数必须大于0",
        ),
        (
            restricted_stock_income,
            ("13.06", "18", 1, 1, "-1"),
            "员工实际出资总额不得为负数",
        ),
        (annual_tax, ("not-a-number",), "应纳税所得额必须是有效数字"),
    ],
)
def test_invalid_parameters_raise_clear_value_error(function, args, message):
    with pytest.raises(ValueError, match=message):
        function(*args)


def test_negative_income_is_zero_with_warning():
    option_result = option_income("10.00", "14.56", 50000)
    restricted_result = restricted_stock_income("5", "5", 100, 100, "1000")

    assert option_result[0] == Decimal("0.00")
    assert option_result[1] == ["股票期权所得为负，已按0处理"]
    assert restricted_result[0] == Decimal("0.00")
    assert restricted_result[1] == ["限制性股票所得为负，已按0处理"]


@pytest.mark.parametrize(
    ("value", "months", "expected"),
    [
        (date(2024, 2, 29), 36, date(2027, 2, 28)),
        (date(2024, 1, 31), 1, date(2024, 2, 29)),
        (date(2024, 4, 30), 1, date(2024, 5, 31)),
        (date(2024, 1, 30), 1, date(2024, 2, 29)),
    ],
)
def test_add_months_handles_month_end(value, months, expected):
    assert add_months(value, months) == expected


def test_installment_status_requires_event_or_tax_batch_link():
    with pytest.raises(ValueError, match="必须关联事件编号或税额批次"):
        installment_status(
            event_id="E1",
            incremental_tax="1000",
            payments=[{"payment_date": date(2026, 1, 1), "amount": "100"}],
            deadline=date(2027, 1, 1),
            tax_batch_id="E1-TAX",
        )


def test_installment_status_flags_overpayment_and_late_payment():
    result = installment_status(
        event_id="E1",
        incremental_tax="1000",
        payments=[
            {
                "event_id": "E1",
                "tax_batch_id": "E1-TAX",
                "payment_date": date(2027, 1, 2),
                "amount": "1100",
            }
        ],
        deadline=date(2027, 1, 1),
        tax_batch_id="E1-TAX",
        as_of_date=date(2027, 1, 2),
    )

    assert result["paid"] == Decimal("1100.00")
    assert result["remaining"] == Decimal("0.00")
    assert result["is_overpaid"] is True
    assert result["is_overdue"] is True
    assert "事件E1累计缴税超过本批次新增税额100.00元" in result["warnings"]
    assert "事件E1存在逾期缴税" in result["warnings"]


def test_installment_status_flags_departure_with_unpaid_tax():
    result = installment_status(
        event_id="E2",
        incremental_tax="1000",
        payments=[
            {
                "event_id": "E2",
                "payment_date": date(2026, 5, 1),
                "amount": "400",
            }
        ],
        deadline=date(2028, 6, 30),
        departure_date=date(2026, 6, 30),
        tax_batch_id="E2-TAX",
    )

    assert result["remaining"] == Decimal("600.00")
    assert result["departure_unpaid"] is True
    assert "事件E2离职前仍有未缴税额600.00元" in result["warnings"]


def test_installment_status_flags_unpaid_balance_after_deadline():
    result = installment_status(
        event_id="E3",
        incremental_tax="1000",
        payments=[],
        deadline=date(2027, 1, 31),
        as_of_date=date(2027, 2, 1),
        tax_batch_id="E3-TAX",
    )

    assert result["remaining"] == Decimal("1000.00")
    assert result["is_overdue"] is True
    assert "事件E3截至2027-02-01已逾期且仍有未缴税额1000.00元" in result["warnings"]


def test_installment_status_excludes_payment_with_only_wrong_tax_batch():
    result = installment_status(
        event_id="E4",
        tax_batch_id="E4-TAX",
        incremental_tax="1000",
        payments=[
            {
                "tax_batch_id": "OTHER-TAX",
                "payment_date": date(2026, 1, 1),
                "amount": "400",
            }
        ],
        deadline=date(2028, 1, 1),
    )

    assert result["paid"] == Decimal("0.00")
    assert result["remaining"] == Decimal("1000.00")
    assert result["has_invalid_payments"] is True
    assert result["invalid_payment_count"] == 1
    assert "第1笔缴税关联税额批次OTHER-TAX与目标批次E4-TAX不一致，未计入" in result["warnings"]


def test_installment_status_excludes_matching_event_with_wrong_tax_batch():
    result = installment_status(
        event_id="E5",
        tax_batch_id="E5-TAX",
        incremental_tax="1000",
        payments=[
            {
                "event_id": "E5",
                "tax_batch_id": "OTHER-TAX",
                "payment_date": date(2026, 1, 1),
                "amount": "600",
            }
        ],
        deadline=date(2028, 1, 1),
    )

    assert result["paid"] == Decimal("0.00")
    assert result["remaining"] == Decimal("1000.00")
    assert result["has_invalid_payments"] is True
    assert result["invalid_payment_count"] == 1
    assert "第1笔缴税关联税额批次OTHER-TAX与目标批次E5-TAX不一致，未计入" in result["warnings"]


def test_installment_status_counts_payment_when_event_and_batch_both_match():
    result = installment_status(
        event_id="E6",
        tax_batch_id="E6-TAX",
        incremental_tax="1000",
        payments=[
            {
                "event_id": "E6",
                "tax_batch_id": "E6-TAX",
                "payment_date": date(2026, 1, 1),
                "amount": "600",
            }
        ],
        deadline=date(2028, 1, 1),
    )

    assert result["paid"] == Decimal("600.00")
    assert result["remaining"] == Decimal("400.00")
    assert result["has_invalid_payments"] is False
    assert result["invalid_payment_count"] == 0
    assert result["warnings"] == []


def test_installment_status_excludes_future_payment_from_historical_balance():
    result = installment_status(
        event_id="E7",
        tax_batch_id="E7-TAX",
        incremental_tax="1000",
        payments=[
            {
                "event_id": "E7",
                "tax_batch_id": "E7-TAX",
                "payment_date": date(2026, 6, 30),
                "amount": "400",
            },
            {
                "event_id": "E7",
                "tax_batch_id": "E7-TAX",
                "payment_date": date(2026, 7, 1),
                "amount": "600",
            },
        ],
        deadline=date(2028, 1, 1),
        as_of_date=date(2026, 6, 30),
    )

    assert result["paid"] == Decimal("400.00")
    assert result["remaining"] == Decimal("600.00")
    assert result["is_overpaid"] is False


def test_installment_status_counts_payment_on_departure_date():
    result = installment_status(
        event_id="E8",
        tax_batch_id="E8-TAX",
        incremental_tax="1000",
        payments=[
            {
                "event_id": "E8",
                "tax_batch_id": "E8-TAX",
                "payment_date": date(2026, 6, 30),
                "amount": "1000",
            }
        ],
        deadline=date(2028, 1, 1),
        departure_date=date(2026, 6, 30),
        as_of_date=date(2026, 6, 30),
    )

    assert result["paid"] == Decimal("1000.00")
    assert result["remaining"] == Decimal("0.00")
    assert result["departure_unpaid"] is False


@pytest.mark.parametrize(
    ("event_id", "tax_batch_id", "message"),
    [
        (None, "E9-TAX", "事件编号不能为空"),
        ("   ", "E9-TAX", "事件编号不能为空"),
        ("E9", None, "税额批次编号不能为空"),
        ("E9", "   ", "税额批次编号不能为空"),
    ],
)
def test_installment_status_rejects_empty_target_identifiers(
    event_id, tax_batch_id, message
):
    with pytest.raises(ValueError, match=message):
        installment_status(
            event_id=event_id,
            tax_batch_id=tax_batch_id,
            incremental_tax="1000",
            payments=[],
            deadline=date(2028, 1, 1),
        )
