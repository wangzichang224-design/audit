from __future__ import annotations

from datetime import datetime
from pathlib import Path

from auditpaper_agent.config import get_settings
from auditpaper_agent.discovery import CashMaterialsDiscovery, discover_cash_materials
from auditpaper_agent.logic.cash import run_cash_workflow
from auditpaper_agent.sensing.cash import ingest_cash_case
from auditpaper_agent.trace import AuditTracer


def run_cash_wizard() -> None:
    """Auditor-friendly interactive CLI for the cash workpaper flow."""
    try:
        from rich.console import Console
        from rich.panel import Panel
        from rich.prompt import Confirm, Prompt
        from rich.table import Table
    except Exception as exc:  # pragma: no cover - dependency/environment guard
        raise RuntimeError("Interactive wizard requires rich. Install project dependencies first.") from exc

    console = Console()
    settings = get_settings()
    console.print(Panel.fit("AuditPaper-Agent 货币资金底稿向导", subtitle="不用记命令，按提示选择文件即可"))

    materials_dir = _ask_path(Prompt, "资料文件夹路径（把 C 底稿需要的材料放在同一个文件夹）", "", must_exist=True)
    discovery = discover_cash_materials(materials_dir)
    _print_discovery(console, Table, discovery)

    default_case_dir = f"run_out\\cash_{datetime.now():%Y%m%d_%H%M%S}"
    case_dir = _ask_path(Prompt, "案件输出目录", default_case_dir, must_exist=False)
    trial_balance = str(discovery.trial_balance) if discovery.trial_balance else _ask_file_path(Prompt, "试算平衡表/TB 文件路径", "")
    journal = str(discovery.journal) if discovery.journal else _ask_file_path(Prompt, "银行存款日记账/序时账 文件路径", "")
    bank_statement = str(discovery.bank_statement) if discovery.bank_statement else _ask_optional_path(Prompt, "银行回函/对账单 文件或目录路径（可空）")

    provider = Prompt.ask(
        "OCR 方式",
        choices=["pdf-text", "textin", "qwen-ocr", "stub"],
        default="pdf-text",
    )
    if provider == "textin":
        console.print("[yellow]TextIn 会调用付费 API。原生 PDF 建议先用 pdf-text，扫描件/图片再用 TextIn。[/yellow]")
        if not Confirm.ask("确认本次使用 TextIn 解析？", default=False):
            provider = "pdf-text"
    if provider == "qwen-ocr":
        console.print("[yellow]Qwen OCR 会调用多模态模型 API。[/yellow]")
        if not Confirm.ask("确认本次使用 Qwen OCR？", default=False):
            provider = "pdf-text"

    client_name = Prompt.ask("客户名称（可空，能从材料读取则自动识别）", default="")
    period_end = Prompt.ask("期末日期 YYYY-MM-DD（可空，能从材料读取则自动识别）", default="")
    template = str(discovery.template) if discovery.template else _ask_file_path(Prompt, "底稿模板 Excel 路径", "")
    default_output = str(Path(case_dir) / "filled_workpaper.xlsx")
    output = _ask_path(Prompt, "输出底稿路径", default_output, must_exist=False)
    use_reasoning = Confirm.ask("是否用 DeepSeek 优化审计发现文字？", default=False)
    show_trace = Confirm.ask("是否显示详细审计轨迹？", default=True)

    tracer = AuditTracer(enabled=show_trace)
    console.rule("开始处理")
    package = ingest_cash_case(
        case_dir=case_dir,
        trial_balance_path=trial_balance,
        journal_path=journal,
        bank_statement_path=bank_statement,
        ocr_provider=provider,
        client_name=client_name or None,
        period_end=period_end or None,
        tracer=tracer,
    )
    result = run_cash_workflow(
        case_dir=case_dir,
        template_path=template,
        output_path=output,
        tracer=tracer,
        use_reasoning=use_reasoning,
    )

    table = Table(title="处理完成")
    table.add_column("项目")
    table.add_column("结果")
    table.add_row("客户", package.meta.client_name)
    table.add_row("TB 行数", str(len(package.trial_balance)))
    table.add_row("银行日记账行数", str(len(package.bank_journal)))
    table.add_row("回函/对账资料行数", str(len(package.bank_confirmations)))
    table.add_row("审计发现", str(len(result.findings)))
    table.add_row("标准材料包", str(Path(case_dir) / "case_package.json"))
    table.add_row("审计发现 JSON", result.findings_path)
    table.add_row("写入计划 JSON", result.write_plan_path)
    table.add_row("证据链 JSON", result.provenance_path or "")
    table.add_row("填好底稿", result.output_path or "")
    console.print(table)

    if settings.ocr_provider == "textin" and provider == "pdf-text":
        console.print("[dim]提示：当前 .env 默认 OCR 是 TextIn，但本向导本次选择了免费 pdf-text。[/dim]")


def _ask_path(prompt_cls, label: str, default: str, must_exist: bool) -> str:
    while True:
        value = _clean_path(prompt_cls.ask(label, default=default))
        if not value:
            print("路径不能为空。")
            continue
        path = Path(value)
        if must_exist and not path.exists():
            print(f"路径不存在：{path}")
            continue
        return str(path)


def _ask_optional_path(prompt_cls, label: str) -> str | None:
    value = _clean_path(prompt_cls.ask(label, default=""))
    if not value:
        return None
    path = Path(value)
    while not path.exists():
        print(f"路径不存在：{path}")
        value = _clean_path(prompt_cls.ask(label, default=""))
        if not value:
            return None
        path = Path(value)
    return str(path)


def _ask_file_path(prompt_cls, label: str, default: str) -> str:
    while True:
        value = _clean_path(prompt_cls.ask(label, default=default))
        if not value:
            print("路径不能为空。")
            continue
        path = Path(value)
        if not path.exists():
            print(f"路径不存在：{path}")
            continue
        if path.is_dir():
            print(f"这里需要选择具体文件，不是文件夹：{path}")
            continue
        return str(path)


def _clean_path(value: str) -> str:
    return str(value or "").strip().strip('"').strip("'")


def _print_discovery(console, table_cls, discovery: CashMaterialsDiscovery) -> None:
    table = table_cls(title="自动识别资料")
    table.add_column("资料")
    table.add_column("识别结果")
    table.add_row("试算平衡表/TB", str(discovery.trial_balance or "未识别"))
    table.add_row("银行日记账/序时账", str(discovery.journal or "未识别"))
    table.add_row("银行回函/对账单", str(discovery.bank_statement or "未识别，可稍后手动补充"))
    table.add_row("底稿模板", str(discovery.template or "未识别"))
    console.print(table)
