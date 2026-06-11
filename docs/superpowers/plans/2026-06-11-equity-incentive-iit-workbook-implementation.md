# 广电计量股权激励个人所得税工作簿 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 依据广电计量正式公告和现行个人所得税政策，生成可批量管理、年度合并计税、自定义36个月缴税及单人查询的可靠 Excel 工作簿。

**Architecture:** 先建立可追溯的公告与政策参数清单，再用独立 Python 计算模块固化税务计算和校验规则，最后由工作簿生成器创建公式、验证、样式和保护。测试既验证 Python 计算结果，也直接检查 XLSX 内部结构、公式和交付文件。

**Tech Stack:** Python 3、openpyxl、pytest、Excel XLSX、公开公告及税务机关官方资料。

---

## 文件结构

- `research/公告及政策依据.md`：公告参数、政策口径和来源链接。
- `src/tax_model.py`：股权激励所得、年度合并税额和期限校验的纯计算函数。
- `src/build_workbook.py`：创建并保存新版 Excel。
- `tests/test_tax_model.py`：计算逻辑单元测试。
- `tests/test_workbook.py`：工作簿结构、公式、验证、保护和错误扫描测试。
- `复核说明.md`：原表问题、修正内容、政策依据和测试结果。
- `广电计量股权激励个人所得税计算及分期台账-2026版.xlsx`：最终交付文件。

### Task 1: 核实公告参数与税务政策

**Files:**
- Create: `research/公告及政策依据.md`

- [ ] **Step 1: 检索公司正式公告**

从巨潮资讯或深交所获取广电计量最近三年与2023年股票期权及限制性股票激励计划有关的正式公告，至少核实授予日、登记完成日、上市日、行权价、授予价、授予数量、激励人数和后续调整。

- [ ] **Step 2: 核实股票登记日收盘价**

从深交所、巨潮资讯引用的行情资料或可追溯市场数据核实2024年8月27日广电计量收盘价。记录数值、日期和来源；不能可靠核实时明确写为“待人工录入”，不得猜值。

- [ ] **Step 3: 核实现行税务口径**

使用财政部、税务总局及主管税务机关官方资料核实：

- 股票期权应纳税所得额；
- 限制性股票应纳税所得额；
- 同一纳税年度多次股权激励合并计税；
- 单独计税政策有效期；
- 境内上市公司36个月缴税政策、备案条件和离职处理。

- [ ] **Step 4: 写入证据清单**

每条结论写明文件名称、文号、发布日期、链接、采用字段和适用说明。不得只写二手解读。

- [ ] **Step 5: 自检**

检查所有数值均有来源，政策有效期均以2026年6月11日为基准，没有将授予价当作登记日收盘价。

- [ ] **Step 6: 提交**

```powershell
git add research/公告及政策依据.md
git commit -m "docs: verify equity incentive announcements and tax rules"
```

### Task 2: 实现并测试税务计算模型

**Files:**
- Create: `src/tax_model.py`
- Create: `tests/test_tax_model.py`

- [ ] **Step 1: 编写失败测试**

覆盖以下具体场景：

- 5万份期权、行权价14.56元、行权日价17.80元，所得额162,000元；
- 限制性股票公式必须分别使用登记日价、解禁日价、解禁数量、总获授数量和实际出资；
- 同一员工同年多次事件合并后只适用一次税率和速算扣除数；
- 跨年度事件分别计税；
- 负所得按零归集并返回警告；
- 36个月截止日正确处理月末；
- 超额缴税、逾期和离职未缴清返回异常。

- [ ] **Step 2: 运行测试确认失败**

```powershell
& 'C:\Users\lucheng\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m pytest tests/test_tax_model.py -v
```

Expected: 因 `src.tax_model` 尚不存在而失败。

- [ ] **Step 3: 实现最小计算模块**

实现：

```python
def option_income(market_price: float, exercise_price: float, quantity: float) -> tuple[float, list[str]]
def restricted_stock_income(registration_price: float, unlock_price: float, unlock_quantity: float, total_granted: float, total_paid: float) -> tuple[float, list[str]]
def annual_tax(taxable_income: float) -> dict[str, float]
def add_months(value: date, months: int) -> date
def installment_status(tax_due: float, payments: list[dict], deadline: date, departure_date: date | None = None) -> dict
```

使用 `Decimal` 进行金额计算并按分四舍五入。参数缺失或数量非法时抛出明确异常。

- [ ] **Step 4: 运行测试确认通过**

```powershell
& 'C:\Users\lucheng\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m pytest tests/test_tax_model.py -v
```

Expected: 全部通过。

- [ ] **Step 5: 提交**

```powershell
git add src/tax_model.py tests/test_tax_model.py
git commit -m "feat: add equity incentive tax calculation model"
```

### Task 3: 生成新版 Excel

**Files:**
- Create: `src/build_workbook.py`
- Create: `tests/test_workbook.py`
- Create: `广电计量股权激励个人所得税计算及分期台账-2026版.xlsx`

- [ ] **Step 1: 编写失败测试**

测试工作簿包含且仅包含以下业务工作表：

```text
使用说明
计划参数
激励事件明细
年度计税汇总
分期缴税台账
单人测算
税率及规则
政策及公告
```

同时检查：

- 原始工作簿未被修改；
- 事件明细包含设计文档规定字段和至少500行可录入空间；
- 计划、权益类型和员工编号有下拉验证；
- 公式列已填充到预留行；
- 年度汇总按员工编号和年度合并；
- 分期台账允许自定义日期及金额；
- 税率表和公式单元格受保护；
- 冻结窗格、筛选、打印设置存在；
- 不存在无说明隐藏试算表；
- 文件内公式不包含 `#REF!`、`#DIV/0!`、`#VALUE!` 或 `#NAME?`。

- [ ] **Step 2: 运行测试确认失败**

```powershell
& 'C:\Users\lucheng\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m pytest tests/test_workbook.py -v
```

Expected: 因生成器和目标工作簿不存在而失败。

- [ ] **Step 3: 实现工作簿生成器**

生成器应：

- 从 `research/公告及政策依据.md` 已核实参数建立计划参数表；
- 使用 Excel 兼容公式，避免依赖新版动态数组函数；
- 使用蓝色表示输入、灰色表示内置参数、绿色表示公式、红色表示异常；
- 用 Excel 表格、数据验证、条件格式、冻结窗格和保护实现可维护模板；
- 加入旧表5万份期权及5万股限制性股票演示案例，员工编号使用 `DEMO001`；
- 设置自动重算模式；
- 保存到项目根目录，绝不覆盖原始文件。

- [ ] **Step 4: 运行生成器**

```powershell
& 'C:\Users\lucheng\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' src/build_workbook.py
```

Expected: 生成 `广电计量股权激励个人所得税计算及分期台账-2026版.xlsx`。

- [ ] **Step 5: 运行测试确认通过**

```powershell
& 'C:\Users\lucheng\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m pytest tests/test_workbook.py -v
```

Expected: 全部通过。

- [ ] **Step 6: 提交**

```powershell
git add src/build_workbook.py tests/test_workbook.py 广电计量股权激励个人所得税计算及分期台账-2026版.xlsx
git commit -m "feat: build equity incentive tax workbook"
```

### Task 4: 复核、实机重算与交付说明

**Files:**
- Create: `复核说明.md`
- Modify: `广电计量股权激励个人所得税计算及分期台账-2026版.xlsx`

- [ ] **Step 1: 全量自动测试**

```powershell
& 'C:\Users\lucheng\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m pytest -v
```

Expected: 全部通过。

- [ ] **Step 2: Excel 实机打开并重算**

优先使用本机 Excel COM 以只读方式打开新版文件、执行完整重算并另存。若 COM 不可用，记录限制并用 LibreOffice headless 重算；两者均不可用时不得声称已实机重算。

- [ ] **Step 3: 重算后扫描**

重新读取公式和值，检查：

- 所有演示结果与 Python 手工计算一致；
- 没有公式错误；
- 保护、验证、筛选和冻结窗格仍存在；
- 原始文件哈希未变化。

- [ ] **Step 4: 编写复核说明**

列明：

- 原表七项结构和计算问题；
- 新版工作表及使用方法；
- 公告参数和政策来源；
- 演示数据复算结果；
- 自动测试、实机重算及残余限制。

- [ ] **Step 5: 提交**

```powershell
git add 复核说明.md 广电计量股权激励个人所得税计算及分期台账-2026版.xlsx
git commit -m "docs: add workbook review and verification results"
```

### Task 5: 最终质量审查

**Files:**
- Review: all project files

- [ ] **Step 1: 规范符合性审查**

逐项核对设计文档、`AGENTS.md` 和本计划，确认无遗漏、无超范围模块。

- [ ] **Step 2: 代码与工作簿质量审查**

检查计算精度、公式引用、参数来源、保护策略、Excel兼容性和可维护性。

- [ ] **Step 3: 最终验证**

```powershell
git status --short
& 'C:\Users\lucheng\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m pytest -v
```

Expected: 仅原始工作簿保持未跟踪或明确排除；测试全部通过。
