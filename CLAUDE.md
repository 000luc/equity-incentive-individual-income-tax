# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

广电计量（002967）2023年股票期权与限制性股票激励计划——个人所得税计算及分期缴税台账。

生成 Excel 工作簿，支持：事件明细录入 → 年度合并计税 → 36个月分期缴税跟踪。工作簿包含8个sheet：使用说明、计划参数、激励事件明细、年度计税汇总、分期缴税台账、单人测算、税率及规则、政策及公告。

## 运行与测试

```bash
# 生成新版工作簿
py src/build_workbook.py

# 运行全部测试
pytest

# 运行单个测试文件
pytest tests/test_tax_model.py

# 运行单个测试
pytest tests/test_tax_model.py::test_option_income_for_50000_options -v
```

依赖：`openpyxl`、`pytest`。

## 代码结构

- **`src/tax_model.py`** — 计税核心逻辑（纯 Python，无外部依赖）。股票期权所得、限制性股票所得、年度累计计税、分期状态判断。
- **`src/build_workbook.py`** — 用 openpyxl 构建完整 Excel 工作簿，含公式、数据验证、条件格式、颜色编码和保护。
- **`tests/test_tax_model.py`** — 计税模型的 pytest 测试，覆盖边界值、异常、批量事件、分期缴税等场景。
- **`tests/test_workbook.py`** — 工作簿集成测试，构建后校验 sheet 结构、公式、数据验证、样式和保护设置。
- **`research/公告及政策依据.md`** — 公司公告参数、政策原文、行权价格调整链条、自检结论。模板参数和计算规则的第一手来源。

## 关键口径

- 同一员工同一纳税年度多次股权激励合并单独计税（财政部税务总局公告2023年第25号）。
- 限制性股票所得公式：（登记日收盘价＋解禁日收盘价）÷2×本批解禁数量－个人实际出资×本批数量÷个人获授总数（国税函〔2009〕461号）。分母必须是员工个人口径，不得使用公司802万股。
- 每笔事件新增税额从该事件行权日/解禁日起单独计算36个月（财政部税务总局公告2024年第2号）。
- 股票期权行权价格按事件日期自动匹配：14.56→14.32→14.18→14.04→13.85元/份。

## 项目文件

- 输出工作簿：`广电计量股权激励个人所得税计算及分期台账-2026版.xlsx`
- 原始参考表：`股权激励个人所得税计算表-V2-2026.06.11 17点.xlsx`
- 设计文档：`2026-06-11-equity-incentive-iit-workbook-design.md`
