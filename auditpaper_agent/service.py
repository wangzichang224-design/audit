from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from auditpaper_agent.discovery import CashMaterialSet, resolve_cash_materials
from auditpaper_agent.logic.cash import run_cash_workflow
from auditpaper_agent.sensing.cash import ingest_cash_case
from auditpaper_agent.trace import AuditTracer, TraceEvent
from auditpaper_agent.utils import safe_filename


@dataclass(frozen=True)
class CashMaterialsSummary:
    materials_dir: Path
    trial_balance: Path | None = None
    journal: Path | None = None
    bank_statement: Path | None = None
    template: Path | None = None
    confidence: float = 0.0
    candidate_sets: list[CashMaterialSet] = field(default_factory=list)
    agent_used: bool = False
    agent_reason: str = ""
    missing_required: list[str] = field(default_factory=list)

    @property
    def is_ready(self) -> bool:
        return not self.missing_required


@dataclass(frozen=True)
class AutoCashRunResult:
    success: bool
    discovery: CashMaterialsSummary
    case_dir: Path | None = None
    output_path: Path | None = None
    client_name: str = ""
    period_end: str | None = None
    findings_count: int = 0
    write_commands_count: int = 0
    provenance_count: int = 0
    artifacts: dict[str, Path] = field(default_factory=dict)
    trace_events: list[TraceEvent] = field(default_factory=list)
    error: str = ""


def clean_materials_path(value: str | Path) -> Path:
    """Normalize a local folder path pasted from chat, Explorer, or quotes."""
    text = str(value).strip().strip("\ufeff").strip()
    text = text.strip("`").strip().strip('"').strip("'").strip()
    if text.lower().startswith("file://"):
        text = text[7:]
    return Path(text).expanduser()


def inspect_cash_materials(materials_dir: str | Path) -> CashMaterialsSummary:
    path = clean_materials_path(materials_dir)
    if not path.is_dir():
        return CashMaterialsSummary(
            materials_dir=path,
            missing_required=[f"资料文件夹不存在：{path}"],
        )

    discovery = resolve_cash_materials(path, use_agent=True)
    return CashMaterialsSummary(
        materials_dir=discovery.materials_dir,
        trial_balance=discovery.trial_balance,
        journal=discovery.journal,
        bank_statement=discovery.bank_statement,
        template=discovery.template,
        confidence=discovery.confidence,
        candidate_sets=discovery.candidate_sets,
        agent_used=discovery.agent_used,
        agent_reason=discovery.agent_reason,
        missing_required=discovery.missing_required,
    )


def run_auto_cash_case(
    materials_dir: str | Path,
    ocr_provider: str = "pdf-text",
    use_reasoning: bool = False,
    use_agent: bool = True,
    case_dir: str | Path | None = None,
    output_path: str | Path | None = None,
) -> AutoCashRunResult:
    path = clean_materials_path(materials_dir)
    if not path.is_dir():
        discovery = CashMaterialsSummary(
            materials_dir=path,
            missing_required=[f"资料文件夹不存在：{path}"],
        )
        return AutoCashRunResult(success=False, discovery=discovery, error="资料文件夹缺少必要文件。")
    raw_discovery = resolve_cash_materials(path, use_agent=use_agent)
    discovery = CashMaterialsSummary(
        materials_dir=raw_discovery.materials_dir,
        trial_balance=raw_discovery.trial_balance,
        journal=raw_discovery.journal,
        bank_statement=raw_discovery.bank_statement,
        template=raw_discovery.template,
        confidence=raw_discovery.confidence,
        candidate_sets=raw_discovery.candidate_sets,
        agent_used=raw_discovery.agent_used,
        agent_reason=raw_discovery.agent_reason,
        missing_required=raw_discovery.missing_required,
    )
    if not discovery.is_ready:
        return AutoCashRunResult(success=False, discovery=discovery, error="资料文件夹缺少必要文件。")

    run_root = Path(case_dir) if case_dir else _default_case_dir(discovery.materials_dir)
    output = Path(output_path) if output_path else run_root / "filled_workpaper.xlsx"
    tracer = AuditTracer(enabled=False)

    try:
        package = ingest_cash_case(
            case_dir=run_root,
            trial_balance_path=discovery.trial_balance,
            journal_path=discovery.journal,
            bank_statement_path=discovery.bank_statement,
            ocr_provider=ocr_provider,
            tracer=tracer,
        )
        workflow_result = run_cash_workflow(
            case_dir=run_root,
            template_path=discovery.template,
            output_path=output,
            tracer=tracer,
            use_reasoning=use_reasoning,
        )
    except Exception as exc:
        return AutoCashRunResult(
            success=False,
            discovery=discovery,
            case_dir=run_root,
            output_path=None,
            error=f"资料解析或底稿生成失败：{exc}",
            trace_events=list(tracer.events),
        )

    artifacts = {
        "case_package": run_root / "case_package.json",
        "audit_findings": Path(workflow_result.findings_path),
        "write_plan": Path(workflow_result.write_plan_path),
        "filled_workpaper": output,
    }
    if workflow_result.provenance_path:
        artifacts["provenance"] = Path(workflow_result.provenance_path)

    return AutoCashRunResult(
        success=True,
        discovery=discovery,
        case_dir=run_root,
        output_path=output,
        client_name=package.meta.client_name,
        period_end=package.meta.period_end.isoformat() if package.meta.period_end else None,
        findings_count=len(workflow_result.findings),
        write_commands_count=_json_list_count(artifacts["write_plan"], "commands"),
        provenance_count=_json_list_count(artifacts.get("provenance"), None),
        artifacts=artifacts,
        trace_events=list(tracer.events),
    )


def _default_case_dir(materials_dir: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path("run_out") / f"{safe_filename(materials_dir.name)}_cash_{stamp}"


def _json_list_count(path: Path | None, key: str | None) -> int:
    if path is None or not path.exists():
        return 0
    data: Any = json.loads(path.read_text(encoding="utf-8"))
    if key is None:
        return len(data) if isinstance(data, list) else 0
    values = data.get(key, []) if isinstance(data, dict) else []
    return len(values) if isinstance(values, list) else 0
