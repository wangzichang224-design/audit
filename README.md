# AuditPaper-Agent 本地审计底稿助手

一句话：在你自己的电脑上打开一个网页，把客户资料包所在文件夹路径粘贴进去，系统先做资料诊断，确认无误后自动生成可下载、可复核的 Excel 审计底稿。

这个项目当前定位是“本地审计工作台 / 公开 Beta 原型”，不是公网 SaaS。它读取的是你电脑上的本地文件夹，所以请在本机运行。

## 审计师快速上手

### 1. 准备客户资料包

把客户资料放在一个文件夹里，推荐结构类似：

```text
客户项目文件夹
├─ c底稿资料
│  ├─ 试算平衡表_TB_2025.xlsx
│  ├─ 序时账_2025.xlsx
│  ├─ 银行函证或对账单文件夹
│  └─ C_货币资金_空白底稿模板.xlsx
├─ 02_master_data
├─ 04_accounting_records
└─ 05_audit_workpapers
```

不要求文件名完全一样。系统会尝试识别 TB、序时账、银行资料、底稿模板；如果资料缺失或列名看不懂，会先给出诊断报告，不会直接生成误导性底稿。

### 2. 启动网页

在 PowerShell 进入项目目录：

```powershell
cd D:\03_AI_Projects\AuditPaper-Agent
python -m streamlit run auditpaper_agent\web_app.py
```

浏览器打开后，页面底部会有一个对话框。把客户资料包路径复制进去，例如：

```text
D:\08_拟真底稿和财务数据\audit_sim_autoparts_2025_v2
```

也可以粘贴更大的父目录，例如：

```text
D:\08_拟真底稿和财务数据
```

系统会先对候选项目打分，尽量选择同一个项目、同一个资料包里的 TB、序时账、函证和模板，避免把不同项目的文件混在一起。

### 3. 先看诊断，再生成

网页会先显示：

- 识别到的资料类型
- 实际使用的客户项目或 C 资料包
- 候选资料包评分
- 缺失资料或无法识别的列名
- 当前 API / Provider 状态

只有诊断通过后，才会出现“确认生成底稿”按钮。

### 4. 下载结果

生成后可以下载：

- 单份 Excel 底稿
- 整套 ZIP
- `suite_manifest.json`
- 关键 JSON 过程文件

默认输出目录在项目内的 `run_out/`，每次生成都会带时间戳，避免覆盖历史结果。

## 现在能生成什么

当前公开 Beta 重点做深：

- `C`：货币资金
- `D`：交易性金融资产及其他权益工具投资
- `EXP`：费用测试与波动分析

完整项目模式还会生成：

- `A10_风险评估与重要性.xlsx`
- `C_货币资金_2025_华衡汽配.xlsx`
- `D10_交易性金融资产及其他权益工具投资.xlsx`
- `E20_应收账款与ECL.xlsx`
- `EXP10_费用测试与波动分析.xlsx`
- `F10_存货监盘与跌价.xlsx`
- `K10_固定资产与折旧.xlsx`
- `N10_应付账款与SURL.xlsx`
- `U10_收入确认与截止.xlsx`

其中 C/D/EXP 自动化深度最高；A/E/F/K/N/U 当前是企业级拟真 scaffold-ready 底稿，包含来源索引、复核提示和基础结构，不声称已经完成完整审计判断。

## 结果为什么可复核

AuditPaper-Agent 不让模型直接改 Excel。它采用受控写入流程：

- 先解析客户资料
- 再生成结构化审计发现和写入计划
- 最后由 Excel harness 写入指定区域
- 公式单元格受保护
- 每个 agent 写入格都保留来源、用途和 SHA256 provenance

也就是说，AI 可以辅助识别、映射、解释和起草发现，但金额计算、写入边界和来源追踪由本地规则控制。

## API 使用情况

默认路径不调用付费 API。

默认使用：

- Python
- pandas
- openpyxl
- pypdf
- 本地规则和本地 PDF 文本抽取

可选增强：

- DeepSeek-compatible reasoning：用于资料歧义判断、列名映射、审计发现文字优化
- Qwen-compatible vision：用于图片/PDF 扫描件识别
- TextIn OCR：用于表格 OCR

API 只参与识别、映射、解释和文字增强，不直接修改 workbook。

如需配置 API，把 `.env.example` 复制为 `.env`，再填入你自己的 token：

```env
AUDITPAPER_REASONING_PROVIDER=deepseek
AUDITPAPER_REASONING_BASE_URL=https://api.deepseek.com/v1
AUDITPAPER_REASONING_TOKEN=
AUDITPAPER_REASONING_MODEL=deepseek-chat

AUDITPAPER_VISION_PROVIDER=qwen
AUDITPAPER_VISION_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
AUDITPAPER_VISION_TOKEN=
AUDITPAPER_VISION_MODEL=qwen-vl-max

AUDITPAPER_OCR_PROVIDER=pdf-text
AUDITPAPER_TEXTIN_APP_ID=
AUDITPAPER_TEXTIN_SECRET_CODE=
AUDITPAPER_TEXTIN_ENDPOINT=
```

`.env` 不应提交到 GitHub。

## 公开 Beta 边界

这个项目目前适合：

- 本地演示“粘贴路径生成底稿”
- 相近结构客户资料包测试
- C/D/EXP 公开 Beta 压测
- 审计 AI 工作流、provenance、人工复核机制展示

暂时不适合直接承诺：

- 完全不同客户资料 100% 自动跑通
- 多人同时在线使用
- 公网 SaaS 直接读取用户电脑路径
- 替代审计人员判断
- 覆盖所有审计科目和所有审计程序

如果资料不完整、列名严重不标准、文件混放多个项目，系统应先给诊断报告，而不是硬生成。

## Clean-room 声明

本项目可以只读扫描你本机参考模板的目录结构，用于理解信息架构和科目覆盖，但不会复制、提交或运行依赖任何会计师事务所专有 Excel 文件。

公开仓库和生成结果不得包含：

- 原始事务所模板文件
- 隐藏参数表
- 宏
- 专有 UUID
- 专有文案
- 专有配色和命名痕迹
- 事务所或第三方专有标记

生成的公开版底稿是 AuditPaper-Agent 自有 clean-room 结构。

## 给开发者看的说明

### 安装依赖

```powershell
cd D:\03_AI_Projects\AuditPaper-Agent
python -m pip install -e .[dev]
```

### 运行测试

```powershell
python -m pytest
```

### Python API

生成完整底稿套件：

```python
from auditpaper_agent.suite import run_auto_workpaper_suite

result = run_auto_workpaper_suite(r"D:\...\audit_sim_autoparts_2025_v2")
print(result.output_dir)
print(result.zip_path)
```

诊断客户资料包：

```python
from auditpaper_agent import diagnose_client_package

diagnostics = diagnose_client_package(r"D:\cases\client_package")
print(diagnostics.mode, diagnostics.can_generate, diagnostics.confidence)
```

发现 C 资料包候选：

```python
from auditpaper_agent import discover_cash_material_sets

sets = discover_cash_material_sets(r"D:\cases")
print(sets[0].root, sets[0].confidence)
```

检查 API provider：

```python
from auditpaper_agent import check_provider_health

health = check_provider_health(live=False)
print(health.reasoning.configured, health.vision.configured)
```

运行公开 Beta 压测：

```python
from auditpaper_agent import run_workpaper_stress_suite

stress = run_workpaper_stress_suite(r"D:\cases\beta_fixtures", focus=("C", "D", "EXP"))
print(stress.success_rate)
```

扫描本地参考模板库存：

```python
from auditpaper_agent.workpaper_catalog import inspect_template_inventory

inventory = inspect_template_inventory(r"C:\path\to\reference_templates")
print(inventory.template_count, inventory.mapped_subjects)
```

SAP / 用友 ERP 导出三步流：

```python
from auditpaper_agent.sensing.erp import diagnose_erp_export, import_erp_export
from auditpaper_agent.suite import run_erp_workpaper_suite

manifest = diagnose_erp_export(r"D:\cases\sap_exports", provider="sap")
print(manifest.blocking_issues)

# 人工复核 mapping_manifest.json 后，再确认导入。
import_result = import_erp_export(
    r"D:\cases\sap_exports",
    r"run_out\sap_case",
    provider="sap",
    confirm_mapping=True,
)

suite = run_erp_workpaper_suite(r"run_out\sap_case")
print(suite.zip_path)
```

### CLI

审计师优先用网页。需要命令行时，可以使用：

```powershell
python -m auditpaper_agent.cli wizard cash
```

全自动 C 底稿：

```powershell
python -m auditpaper_agent.cli auto cash `
  --materials-dir path\to\cash_case_folder `
  --trace
```

分步 ingest / run：

```powershell
auditpaper ingest cash `
  --case-dir cases\cash_demo `
  --trial-balance path\to\试算平衡表.xlsx `
  --journal path\to\银行存款日记账.xlsx `
  --bank-statement path\to\bank_confirmations `
  --ocr-provider pdf-text `
  --trace

auditpaper run cash `
  --case-dir cases\cash_demo `
  --template path\to\C_货币资金_空白底稿模板.xlsx `
  --output output\cash_demo_filled.xlsx `
  --trace
```

ERP 导出诊断、确认导入、生成主循环 clean-room 底稿：

```powershell
auditpaper diagnose erp `
  --export-dir path\to\sap_or_yonyou_exports `
  --provider auto `
  --output-dir run_out\erp_mapping

auditpaper ingest erp `
  --export-dir path\to\sap_or_yonyou_exports `
  --case-dir run_out\erp_case `
  --provider auto `
  --confirm-mapping

auditpaper run suite `
  --erp-case-dir run_out\erp_case `
  --output-dir run_out\erp_suite
```

注意：`ingest erp` 默认只写 `mapping_manifest.json` 并阻断生成；只有人工复核字段映射后显式传入 `--confirm-mapping`，才会写入 `erp_package.json` 并允许生成 A/C/D/E/EXP/F/K/N/U 主循环底稿。

### 项目模块

- `sensing/`：解析外部客户资料
- `logic/`：执行确定性审计检查
- `harness/`：控制 Excel 写入、公式保护和 provenance
- `knowledge/`：保存底稿映射和审计措辞默认值
- `suite/`：从本地项目文件夹生成整套拟真审计底稿
- `sensing/erp.py`：SAP / 用友 ERP 导出识别、字段映射诊断和标准化数据包生成
- `diagnostics.py`：公开 Beta 资料诊断
- `agent.py`：OpenAI-compatible reasoning / vision provider 接入

### 设计规则

- 审计逻辑不直接写 workbook
- Excel 写入必须先形成 `WritePlan` / `WriteCellCommand`
- harness 拒绝写入白名单外区域
- harness 拒绝覆盖公式单元格
- 每个写入必须带来源文件 hash 和定位信息
- 金额计算由确定性 Python 代码完成
- LLM/OCR provider 只能辅助识别、映射、解释和措辞，不能直接改 Excel
