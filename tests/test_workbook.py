from __future__ import annotations

import hashlib
import subprocess
import sys
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest
from openpyxl.utils import get_column_letter

from src.tax_model import (
    add_months,
    annual_tax,
    event_tax_batches,
    installment_status,
    option_income,
    restricted_stock_income,
)


ROOT = Path(__file__).resolve().parents[1]
BUILDER = ROOT / "src" / "build_workbook.py"
OUTPUT = ROOT / "广电计量股权激励个人所得税计算及分期台账-2026版.xlsx"
ORIGINAL = ROOT / "股权激励个人所得税计算表-V2-2026.06.11 17点.xlsx"
ORIGINAL_SHA256 = "A320E140E68AF70EAC909E90502C5BC120882D447CFEB141415AA5FE0063E6A0"
SHEETS = [
    "使用说明",
    "计划参数",
    "激励事件明细",
    "年度计税汇总",
    "分期缴税台账",
    "单人测算",
    "税率及规则",
    "政策及公告",
]
ERROR_TOKENS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?")


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


@pytest.fixture(scope="session")
def workbook():
    assert BUILDER.exists(), "工作簿生成器尚未创建"
    subprocess.run([sys.executable, str(BUILDER)], cwd=ROOT, check=True)
    assert OUTPUT.exists(), "新版工作簿尚未生成"
    return openpyxl.load_workbook(OUTPUT, data_only=False)


def header_map(ws) -> dict[str, int]:
    return {cell.value: cell.column for cell in ws[1] if cell.value}


def formula_cells(ws):
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type == "f":
                yield cell


def test_original_workbook_is_unchanged():
    assert sha256(ORIGINAL) == ORIGINAL_SHA256


def test_workbook_has_exactly_eight_visible_business_sheets(workbook):
    assert workbook.sheetnames == SHEETS
    assert all(workbook[name].sheet_state == "visible" for name in SHEETS)
    assert workbook.calculation.calcMode == "auto"


def test_plan_parameters_include_verified_announcement_facts(workbook):
    ws = workbook["计划参数"]
    values = [cell.value for row in ws.iter_rows() for cell in row if cell.value is not None]
    text = "\n".join(str(value) for value in values)
    assert "2023年股票期权与限制性股票激励计划" in text
    assert "实际授予" in text and "完成登记" in text
    normalized_dates = {
        value.date() if isinstance(value, datetime) else value
        for value in values
        if isinstance(value, (date, datetime))
    }
    assert date(2024, 8, 27) in normalized_dates
    assert 13.06 in values
    for price in (14.56, 14.32, 14.18, 14.04):
        assert price in values
    assert "802万股不得作为个人分母" in text
    assert ws.protection.sheet


def test_event_sheet_has_required_columns_and_500_input_rows(workbook):
    ws = workbook["激励事件明细"]
    headers = header_map(ws)
    required = {
        "事件编号",
        "税额批次",
        "员工编号",
        "姓名",
        "激励计划",
        "权益类型",
        "事件日期",
        "纳税年度",
        "本次行权或解禁数量",
        "行权日或解禁日收盘价",
        "适用行权或授予价格",
        "限制性股票登记日收盘价",
        "员工获授限制性股票总数",
        "员工实际出资总额",
        "原始本次所得",
        "本次应纳税所得额",
        "此前年度累计所得",
        "本次事件后年度累计所得",
        "累计税额",
        "此前累计已确认税额",
        "本次新增税额",
        "事件级36个月截止日",
        "校验",
    }
    assert required <= headers.keys()
    assert ws.max_row >= 501
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref
    assert ws.tables

    formula_headers = {
        "税额批次",
        "纳税年度",
        "适用行权或授予价格",
        "限制性股票登记日收盘价",
        "原始本次所得",
        "本次应纳税所得额",
        "此前年度累计所得",
        "本次事件后年度累计所得",
        "累计税额",
        "此前累计已确认税额",
        "本次新增税额",
        "事件级36个月截止日",
        "校验",
    }
    for header in formula_headers:
        col = get_column_letter(headers[header])
        assert ws[f"{col}501"].data_type == "f", f"{header}未填充到第500条数据行"


def test_demo_rows_and_expected_tax_match_tax_model(workbook):
    ws = workbook["激励事件明细"]
    headers = header_map(ws)
    employee_col = headers["员工编号"]
    demo_rows = [
        row
        for row in range(2, ws.max_row + 1)
        if ws.cell(row, employee_col).value == "DEMO001"
    ]
    assert demo_rows == [2, 3]

    option_result, _ = option_income("17.80", "14.56", "50000")
    restricted_result, _ = restricted_stock_income(
        "13.06", "17.80", "50000", "50000", "434000"
    )
    assert option_result == Decimal("162000.00")
    assert restricted_result == Decimal("337500.00")
    assert annual_tax(option_result)["tax"] == Decimal("15480.00")
    assert annual_tax(option_result + restricted_result)["tax"] == Decimal("96930.00")
    batches = event_tax_batches(
        [
            {
                "event_id": "DEMO-OPT-001",
                "employee_id": "DEMO001",
                "event_date": date(2025, 1, 9),
                "income": option_result,
            },
            {
                "event_id": "DEMO-RS-001",
                "employee_id": "DEMO001",
                "event_date": date(2025, 6, 30),
                "income": restricted_result,
            },
        ]
    )
    assert [item["incremental_tax"] for item in batches] == [
        Decimal("15480.00"),
        Decimal("81450.00"),
    ]
    assert [item["deadline"] for item in batches] == [
        date(2028, 1, 9),
        date(2028, 6, 30),
    ]

    assert ws.cell(2, headers["本次行权或解禁数量"]).value == 50000
    assert ws.cell(2, headers["行权日或解禁日收盘价"]).value == 17.8
    assert ws.cell(3, headers["员工获授限制性股票总数"]).value == 50000
    assert ws.cell(3, headers["员工实际出资总额"]).value == 434000
    assert "仅演示" in ws.cell(2, headers["备注"]).value
    assert "仅演示" in ws.cell(3, headers["备注"]).value

    formulas = {
        header: ws.cell(row, headers[header]).value
        for row in demo_rows
        for header in (
            "本次应纳税所得额",
            "此前年度累计所得",
            "本次事件后年度累计所得",
            "累计税额",
            "此前累计已确认税额",
            "本次新增税额",
            "事件级36个月截止日",
        )
    }
    joined = "\n".join(formulas.values())
    assert "SUMIFS" in joined
    assert "EDATE" in joined
    assert "VLOOKUP" in joined


def test_demo_mirror_covers_cross_year_and_month_end():
    batches = event_tax_batches(
        [
            {
                "event_id": "Y2024",
                "employee_id": "E001",
                "event_date": date(2024, 12, 31),
                "income": Decimal("36000"),
            },
            {
                "event_id": "Y2025",
                "employee_id": "E001",
                "event_date": date(2025, 1, 31),
                "income": Decimal("36000"),
            },
        ]
    )
    assert [item["cumulative_income"] for item in batches] == [
        Decimal("36000.00"),
        Decimal("36000.00"),
    ]
    assert add_months(date(2024, 2, 29), 36) == date(2027, 2, 28)
    assert add_months(date(2025, 1, 31), 36) == date(2028, 1, 31)


def test_event_accumulation_uses_date_and_event_id_stable_order(workbook):
    ws = workbook["激励事件明细"]
    headers = header_map(ws)
    formula = ws.cell(2, headers["此前年度累计所得"]).value
    assert "事件日期" not in formula
    date_col = get_column_letter(headers["事件日期"])
    event_col = get_column_letter(headers["事件编号"])
    assert f"${date_col}$2:${date_col}$501" in formula
    assert f"${event_col}$2:${event_col}$501" in formula
    assert '"<"&' in formula
    assert '"<"&' in formula or '"<="' in formula


def test_installment_ledger_links_event_batch_and_uses_event_deadline(workbook):
    ws = workbook["分期缴税台账"]
    headers = header_map(ws)
    required = {
        "员工编号",
        "姓名",
        "纳税年度",
        "事件编号",
        "新增税额批次",
        "该事件新增应纳税额",
        "该事件纳税义务发生日",
        "该事件最晚缴清日",
        "缴税日期",
        "计划缴税金额",
        "实际缴税金额",
        "累计已缴",
        "剩余税额",
        "是否逾期",
        "关联有效",
        "校验",
    }
    assert required <= headers.keys()
    deadline_formula = ws.cell(2, headers["该事件最晚缴清日"]).value
    validation_formula = ws.cell(2, headers["校验"]).value
    link_formula = ws.cell(2, headers["关联有效"]).value
    paid_formula = ws.cell(2, headers["累计已缴"]).value
    assert "激励事件明细" in deadline_formula
    assert "冲突" in validation_formula
    assert "离职后缴税" in validation_formula
    assert "COUNTIF" in link_formula
    assert "有效" in link_formula
    assert get_column_letter(headers["关联有效"]) in paid_formula
    assert ws.max_row >= 501


def test_event_and_batch_conflict_and_departure_late_payment_mirror():
    conflict = installment_status(
        "E1",
        "1000",
        [
            {
                "event_id": "E1",
                "tax_batch_id": "E2-TAX",
                "payment_date": date(2025, 2, 1),
                "amount": "1000",
            }
        ],
        date(2028, 1, 1),
        tax_batch_id="E1-TAX",
        as_of_date=date(2025, 2, 2),
    )
    assert conflict["paid"] == Decimal("0.00")
    assert conflict["has_invalid_payments"]

    late_departure = installment_status(
        "E1",
        "1000",
        [
            {
                "event_id": "E1",
                "tax_batch_id": "E1-TAX",
                "payment_date": date(2025, 3, 2),
                "amount": "1000",
            }
        ],
        date(2028, 1, 1),
        departure_date=date(2025, 3, 1),
        tax_batch_id="E1-TAX",
        as_of_date=date(2025, 3, 3),
    )
    assert late_departure["paid"] == Decimal("1000.00")
    assert late_departure["departure_unpaid"]


def test_event_prices_reference_plan_parameters_without_hardcoding(workbook):
    ws = workbook["激励事件明细"]
    headers = header_map(ws)
    formulas = "\n".join(
        ws.cell(4, headers[name]).value
        for name in (
            "适用行权或授予价格",
            "限制性股票登记日收盘价",
        )
    )
    assert formulas.count("'计划参数'!") >= 2
    for literal in ("14.56", "14.32", "14.18", "14.04", "13.06", "8.68"):
        assert literal not in formulas


def test_negative_raw_income_is_zero_and_warned(workbook):
    ws = workbook["激励事件明细"]
    headers = header_map(ws)
    raw_formula = ws.cell(4, headers["原始本次所得"]).value
    taxable_formula = ws.cell(4, headers["本次应纳税所得额"]).value
    validation_formula = ws.cell(4, headers["校验"]).value
    assert "MAX(0" not in raw_formula
    assert "MAX(0" in taxable_formula
    assert "原始所得为负" in validation_formula


def test_annual_summary_uses_non_array_helper_key_lookup(workbook):
    ws = workbook["年度计税汇总"]
    headers = header_map(ws)
    formula = ws.cell(2, headers["最早到期事件编号"]).value
    paid_formula = ws.cell(2, headers["已缴金额"]).value
    assert "MATCH(1,(" not in formula
    assert "INDEX" in formula and "MATCH" in formula
    assert "年度到期键" in header_map(workbook["激励事件明细"])
    assert "'分期缴税台账'!$Q$2:$Q$501" in paid_formula
    assert '"有效"' in paid_formula


def test_single_person_sheet_only_references_source_sheets(workbook):
    ws = workbook["单人测算"]
    formulas = "\n".join(cell.value for cell in formula_cells(ws))
    assert "激励事件明细" in formulas
    assert "年度计税汇总" in formulas
    assert "分期缴税台账" in formulas
    assert "VLOOKUP" not in formulas or "税率及规则" not in formulas
    values = [cell.value for row in ws.iter_rows() for cell in row if cell.value]
    for heading in ("事件明细", "分期记录", "批次余额", "到期/逾期提示"):
        assert heading in values
    assert ws["B3"].data_type != "f"
    assert "'分期缴税台账'!$Q$2:$Q$501" in ws["B8"].value
    validation = next(
        item for item in ws.data_validations.dataValidation if "B3" in str(item.sqref)
    )
    assert "单人测算" in validation.formula1


def test_validations_styles_protection_and_print_settings(workbook):
    event_ws = workbook["激励事件明细"]
    ledger_ws = workbook["分期缴税台账"]
    summary_ws = workbook["年度计税汇总"]
    event_validations = event_ws.data_validations.dataValidation
    assert not any(
        "C2:C501" in str(validation.sqref) or "C4:C501" in str(validation.sqref)
        for validation in event_validations
    )
    assert len(ledger_ws.data_validations.dataValidation) >= 3

    for ws in (event_ws, ledger_ws, summary_ws):
        assert ws.freeze_panes
        assert ws.auto_filter.ref
        assert ws.print_area
        assert ws.sheet_properties.pageSetUpPr.fitToPage
        assert ws.protection.sheet

    headers = header_map(event_ws)
    input_cell = event_ws.cell(4, headers["员工编号"])
    formula_cell = event_ws.cell(4, headers["纳税年度"])
    assert input_cell.fill.fgColor.rgb.endswith("DDEBF7")
    assert not input_cell.protection.locked
    assert formula_cell.fill.fgColor.rgb.endswith("E2F0D9")
    assert formula_cell.protection.locked
    assert event_ws.conditional_formatting


def test_no_formula_contains_excel_error_tokens(workbook):
    for ws in workbook.worksheets:
        for cell in formula_cells(ws):
            assert not any(token in cell.value for token in ERROR_TOKENS), (
                ws.title,
                cell.coordinate,
                cell.value,
            )
