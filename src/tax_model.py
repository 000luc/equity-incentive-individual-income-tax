from __future__ import annotations

import calendar
from collections import defaultdict
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any


CENT = Decimal("0.01")
ZERO = Decimal("0.00")
TAX_BRACKETS = (
    (Decimal("36000"), Decimal("0.03"), Decimal("0")),
    (Decimal("144000"), Decimal("0.10"), Decimal("2520")),
    (Decimal("300000"), Decimal("0.20"), Decimal("16920")),
    (Decimal("420000"), Decimal("0.25"), Decimal("31920")),
    (Decimal("660000"), Decimal("0.30"), Decimal("52920")),
    (Decimal("960000"), Decimal("0.35"), Decimal("85920")),
    (None, Decimal("0.45"), Decimal("181920")),
)


def _decimal(value: Any, field_name: str) -> Decimal:
    if isinstance(value, bool) or value is None:
        raise ValueError(f"{field_name}必须是有效数字")
    try:
        result = Decimal(str(value))
    except (InvalidOperation, ValueError):
        raise ValueError(f"{field_name}必须是有效数字") from None
    if not result.is_finite():
        raise ValueError(f"{field_name}必须是有效数字")
    return result


def _money(value: Decimal, field_name: str = "金额") -> Decimal:
    try:
        return value.quantize(CENT, rounding=ROUND_HALF_UP)
    except InvalidOperation:
        raise ValueError(f"{field_name}超出可计算范围") from None


def _identifier(value: Any, field_name: str) -> str:
    if value is None or isinstance(value, bool):
        raise ValueError(f"{field_name}不能为空")
    result = str(value).strip()
    if not result:
        raise ValueError(f"{field_name}不能为空")
    return result


def _optional_identifier(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def option_income(
    market_price: Any, exercise_price: Any, quantity: Any
) -> tuple[Decimal, list[str]]:
    market = _decimal(market_price, "行权日市场价")
    exercise = _decimal(exercise_price, "行权价格")
    qty = _decimal(quantity, "行权数量")
    if market < 0:
        raise ValueError("行权日市场价不得为负数")
    if exercise < 0:
        raise ValueError("行权价格不得为负数")
    if qty <= 0:
        raise ValueError("行权数量必须大于0")

    income = (market - exercise) * qty
    if income < 0:
        return ZERO, ["股票期权所得为负，已按0处理"]
    return _money(income), []


def restricted_stock_income(
    registration_price: Any,
    unlock_price: Any,
    unlock_quantity: Any,
    total_granted: Any,
    total_paid: Any,
) -> tuple[Decimal, list[str]]:
    registration = _decimal(registration_price, "股票登记日市场价")
    unlock = _decimal(unlock_price, "解禁日市场价")
    unlock_qty = _decimal(unlock_quantity, "本批解禁数量")
    employee_total = _decimal(total_granted, "员工实际登记总股数")
    employee_paid = _decimal(total_paid, "员工实际出资总额")

    if registration < 0:
        raise ValueError("股票登记日市场价不得为负数")
    if unlock < 0:
        raise ValueError("解禁日市场价不得为负数")
    if unlock_qty <= 0:
        raise ValueError("本批解禁数量必须大于0")
    if employee_total <= 0:
        raise ValueError("员工实际登记总股数必须大于0")
    if unlock_qty > employee_total:
        raise ValueError("本批解禁数量不得超过员工实际登记总股数")
    if employee_paid < 0:
        raise ValueError("员工实际出资总额不得为负数")

    income = (
        (registration + unlock) / Decimal("2") * unlock_qty
        - employee_paid * unlock_qty / employee_total
    )
    if income < 0:
        return ZERO, ["限制性股票所得为负，已按0处理"]
    return _money(income), []


def annual_tax(taxable_income: Any) -> dict[str, Decimal]:
    income = _decimal(taxable_income, "应纳税所得额")
    if income < 0:
        raise ValueError("应纳税所得额不得为负数")

    for ceiling, rate, quick_deduction in TAX_BRACKETS:
        if ceiling is None or income <= ceiling:
            tax = max(income * rate - quick_deduction, Decimal("0"))
            return {
                "taxable_income": _money(income, "应纳税所得额"),
                "rate": rate,
                "quick_deduction": _money(quick_deduction),
                "tax": _money(tax, "应纳税额"),
            }
    raise RuntimeError("无法匹配税率")


def incremental_event_tax(
    cumulative_income: Any, previous_confirmed_tax: Any
) -> dict[str, Decimal]:
    cumulative = annual_tax(cumulative_income)
    previous = _decimal(previous_confirmed_tax, "此前累计已确认税额")
    if previous < 0:
        raise ValueError("此前累计已确认税额不得为负数")
    previous = _money(previous)
    incremental = cumulative["tax"] - previous
    if incremental < 0:
        raise ValueError("此前累计已确认税额不得超过本次后年度累计税额")
    return {
        "cumulative_tax": cumulative["tax"],
        "previous_confirmed_tax": previous,
        "incremental_tax": _money(incremental),
    }


def add_months(value: date, months: int) -> date:
    if not isinstance(value, date):
        raise ValueError("日期必须是date类型")
    if isinstance(months, bool) or not isinstance(months, int):
        raise ValueError("月数必须是整数")

    source_last_day = calendar.monthrange(value.year, value.month)[1]
    month_index = value.year * 12 + value.month - 1 + months
    target_year, zero_based_month = divmod(month_index, 12)
    target_month = zero_based_month + 1
    target_last_day = calendar.monthrange(target_year, target_month)[1]
    if value.day == source_last_day or value.day > target_last_day:
        target_day = target_last_day
    else:
        target_day = value.day
    return date(target_year, target_month, target_day)


def event_tax_batches(events: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not isinstance(events, list):
        raise ValueError("事件必须是列表")

    normalized = []
    seen_event_ids: set[str] = set()
    for event in events:
        if not isinstance(event, dict):
            raise ValueError("每个事件必须是字典")
        for field in ("event_id", "employee_id", "event_date", "income"):
            if field not in event:
                raise ValueError(f"事件缺少必填字段：{field}")
        event_id = _identifier(event["event_id"], "事件编号")
        employee_id = _identifier(event["employee_id"], "员工编号")
        event_date = event["event_date"]
        if event_id in seen_event_ids:
            raise ValueError(f"事件编号重复：{event_id}")
        if not isinstance(event_date, date):
            raise ValueError(f"事件{event_id}的事件日期必须是date类型")
        income = _decimal(event["income"], f"事件{event_id}所得额")
        warnings = list(event.get("warnings", []))
        if income < 0:
            income = Decimal("0")
            warnings.append(f"事件{event_id}所得为负，已按0处理")
        normalized.append(
            {
                **event,
                "event_id": event_id,
                "employee_id": employee_id,
                "event_date": event_date,
                "income": _money(income),
                "warnings": warnings,
            }
        )
        seen_event_ids.add(event_id)

    normalized.sort(
        key=lambda item: (
            item["employee_id"],
            item["event_date"].year,
            item["event_date"],
            item["event_id"],
        )
    )
    cumulative_income: defaultdict[tuple[str, int], Decimal] = defaultdict(
        lambda: Decimal("0")
    )
    confirmed_tax: defaultdict[tuple[str, int], Decimal] = defaultdict(
        lambda: Decimal("0")
    )
    batches = []
    for event in normalized:
        tax_year = event["event_date"].year
        key = (event["employee_id"], tax_year)
        previous_tax = confirmed_tax[key]
        cumulative_income[key] += event["income"]
        tax_result = incremental_event_tax(cumulative_income[key], previous_tax)
        confirmed_tax[key] = tax_result["cumulative_tax"]
        batches.append(
            {
                **event,
                "tax_year": tax_year,
                "cumulative_income": _money(cumulative_income[key]),
                "cumulative_tax": tax_result["cumulative_tax"],
                "previous_confirmed_tax": tax_result["previous_confirmed_tax"],
                "incremental_tax": tax_result["incremental_tax"],
                "tax_obligation_date": event["event_date"],
                "deadline": add_months(event["event_date"], 36),
            }
        )
    return batches


def installment_status(
    event_id: str,
    incremental_tax: Any,
    payments: list[dict[str, Any]],
    deadline: date,
    departure_date: date | None = None,
    as_of_date: date | None = None,
    tax_batch_id: str | None = None,
) -> dict[str, Any]:
    event_id = _identifier(event_id, "事件编号")
    tax_batch_id = _identifier(tax_batch_id, "税额批次编号")
    tax = _decimal(incremental_tax, "本批次新增税额")
    if tax < 0:
        raise ValueError("本批次新增税额不得为负数")
    tax = _money(tax)
    if not isinstance(payments, list):
        raise ValueError("缴税记录必须是列表")
    if not isinstance(deadline, date):
        raise ValueError("最晚缴清日必须是date类型")
    if departure_date is not None and not isinstance(departure_date, date):
        raise ValueError("离职日期必须是date类型")
    if as_of_date is not None and not isinstance(as_of_date, date):
        raise ValueError("检查日期必须是date类型")
    check_date = as_of_date or date.today()

    paid = Decimal("0")
    paid_by_departure = Decimal("0")
    has_late_payment = False
    invalid_payment_count = 0
    warnings = []
    for index, payment in enumerate(payments, start=1):
        if not isinstance(payment, dict):
            raise ValueError(f"第{index}笔缴税记录必须是字典")
        linked_event = _optional_identifier(payment.get("event_id"))
        linked_batch = _optional_identifier(payment.get("tax_batch_id"))
        if not linked_event and not linked_batch:
            raise ValueError("每笔缴税必须关联事件编号或税额批次")
        mismatch_messages = []
        if linked_event and linked_event != event_id:
            mismatch_messages.append(
                f"关联事件{linked_event}与目标事件{event_id}不一致"
            )
        if linked_batch and linked_batch != tax_batch_id:
            mismatch_messages.append(
                f"关联税额批次{linked_batch}与目标批次{tax_batch_id}不一致"
            )
        if mismatch_messages:
            invalid_payment_count += 1
            warnings.append(f"第{index}笔缴税{'；'.join(mismatch_messages)}，未计入")
            continue
        payment_date = payment.get("payment_date")
        if not isinstance(payment_date, date):
            raise ValueError(f"第{index}笔缴税日期必须是date类型")
        amount = _decimal(payment.get("amount"), f"第{index}笔缴税金额")
        if amount <= 0:
            raise ValueError(f"第{index}笔缴税金额必须大于0")
        if payment_date > check_date:
            continue
        paid += amount
        if payment_date > deadline:
            has_late_payment = True
        if departure_date is None or payment_date <= departure_date:
            paid_by_departure += amount

    paid = _money(paid)
    overpaid = max(paid - tax, Decimal("0"))
    remaining = max(tax - paid, Decimal("0"))
    unpaid_overdue = check_date > deadline and remaining > 0
    departure_remaining = (
        max(tax - _money(paid_by_departure), Decimal("0"))
        if departure_date is not None
        else Decimal("0")
    )
    if overpaid > 0:
        warnings.append(f"事件{event_id}累计缴税超过本批次新增税额{_money(overpaid):.2f}元")
    if has_late_payment:
        warnings.append(f"事件{event_id}存在逾期缴税")
    if unpaid_overdue:
        warnings.append(
            f"事件{event_id}截至{check_date.isoformat()}已逾期且仍有未缴税额"
            f"{_money(remaining):.2f}元"
        )
    if departure_remaining > 0:
        warnings.append(
            f"事件{event_id}离职前仍有未缴税额{_money(departure_remaining):.2f}元"
        )

    return {
        "event_id": event_id,
        "tax_batch_id": tax_batch_id,
        "incremental_tax": tax,
        "paid": paid,
        "remaining": _money(remaining),
        "is_overpaid": overpaid > 0,
        "is_overdue": has_late_payment or unpaid_overdue,
        "departure_unpaid": departure_remaining > 0,
        "has_invalid_payments": invalid_payment_count > 0,
        "invalid_payment_count": invalid_payment_count,
        "warnings": warnings,
    }
