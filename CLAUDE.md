# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

广电计量（002967）2023年股票期权与限制性股票激励计划——个人所得税计算及分期缴税台账。

生成 Excel 工作簿，支持：事件明细录入 → 年度合并计税 → 36个月分期缴税跟踪。工作簿包含8个sheet：使用说明、计划参数、激励事件明细、年度计税汇总、分期缴税台账、单人测算、税率及规则、政策及公告。

## 常用命令

依赖：`openpyxl`、`pytest`。本机稳定 Python 路径为 `C:\Users\lucheng\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe`，系统 `python` 不可用时用该路径。

```powershell
# 生成新版工作簿（输出到项目根目录）
py src/build_workbook.py

# 或指定输出路径
$env:WORKBOOK_OUTPUT="C:\path\to\output.xlsx"; py src/build_workbook.py

# 运行全部测试（约80-90秒，含 Excel COM 验证）
pytest -q

# 运行单个测试文件
pytest tests/test_tax_model.py -v
pytest tests/test_workbook.py -v

# 运行单个测试
pytest tests/test_tax_model.py::test_option_income_for_50000_options -v
```

## 代码架构

### 两层镜像设计

- **`src/tax_model.py`** — 计税核心逻辑（纯 Python，无外部依赖）。
  - 股票期权所得、限制性股票所得、年度累计计税、分期状态判断。
  - 所有金额使用 `Decimal` 并按分四舍五入，校验边界和异常。
  - `event_tax_batches()` 按员工、纳税年度、事件日期、事件编号稳定排序，逐笔计算累计所得、累计税额、本次新增税额和36个月截止日。

- **`src/build_workbook.py`** — 用 openpyxl 构建完整 Excel 工作簿。
  - 将 `tax_model.py` 中的计税规则翻译成 Excel 公式（SUMIFS、VLOOKUP、EDATE、COUNTIFS 等）。
  - 公式不依赖动态数组函数，保持与 Excel 16.0 兼容。
  - 定义名称（员工主数据编号、检查截至日期等）供跨表引用。
  - 通过颜色、数据验证、条件格式、表格、保护和打印设置实现可维护模板。

### 测试层

- **`tests/test_tax_model.py`** — 计税模型的 pytest 测试，覆盖边界值、异常、批量事件、分期缴税等场景。
- **`tests/test_workbook.py`** — 工作簿集成测试，构建后校验 sheet 结构、公式、数据验证、样式、保护设置，并在 Windows 上调用 Excel COM 执行完整重算。

## 关键口径

- 同一员工同一纳税年度多次股权激励合并单独计税（财政部税务总局公告2023年第25号）。
- 限制性股票所得公式：（登记日收盘价＋解禁日收盘价）÷2×本批解禁数量－个人实际出资×本批数量÷个人获授总数（国税函〔2009〕461号）。分母必须是员工个人口径，不得使用公司802万股。
- 每笔事件新增税额从该事件行权日/解禁日起单独计算36个月（财政部税务总局公告2024年第2号）。
- 股票期权行权价格按事件日期自动匹配：14.56→14.32→14.18→14.04元/份。

## 项目文件

- 输出工作簿：`广电计量股权激励个人所得税计算及分期台账-2026版.xlsx`
- 原始参考表：`股权激励个人所得税计算表-V2-2026.06.11 17点.xlsx`（只读，不得修改或加入 Git）
- 设计文档：`2026-06-11-equity-incentive-iit-workbook-design.md`
- 实施计划：`docs/superpowers/plans/2026-06-11-equity-incentive-iit-workbook-implementation.md`
- 项目进度与接手说明：`项目进度.md`

## 工作约束（来自 AGENTS.md）

- 全部回复、文档和工作表说明使用中文。
- 直接执行并交付结果，不停留在方案层。
- 原始文件只读保留，修改后必须另存新文件。
- 税务结论必须有现行有效政策依据，不凭经验猜测。
- 广电计量激励计划参数必须取自公司正式公告，并记录公告名称、日期和链接。
- 行权日、解禁日、股票登记日的市场价格必须注明可靠来源和日期。
- Excel 交付前必须检查公式错误、重复事件、缺失参数、跨年度归集、超额缴税和逾期状态。
