from __future__ import annotations

from pathlib import Path
from uuid import uuid4
import os

import pytest

from auditpaper_agent.logic.cash import run_cash_workflow
from auditpaper_agent.sensing.cash import ingest_cash_case
from auditpaper_agent.trace import AuditTracer


SAMPLE_DIR = Path(r"D:\03_AI_Projects\cross_border_audit_agent\cross_border_audit_agent\benchmarks\materials\cash")
EXTERNAL_TEMPLATE_DIR = Path(os.getenv("AUDITPAPER_EXTERNAL_CASH_TEMPLATE_DIR", ""))


def _external_template() -> Path | None:
    if not str(EXTERNAL_TEMPLATE_DIR) or not EXTERNAL_TEMPLATE_DIR.exists():
        return None
    matches = list(EXTERNAL_TEMPLATE_DIR.glob("*货币资金*.xlsx"))
    return matches[0] if matches else None


def _runtime_dir() -> Path:
    path = Path(__file__).resolve().parents[1] / "test_runtime" / uuid4().hex
    path.mkdir(parents=True, exist_ok=True)
    return path


@pytest.mark.skipif(not SAMPLE_DIR.exists(), reason="external sample cash materials not available")
def test_ingest_external_cash_materials() -> None:
    case_dir = _runtime_dir() / "cash_demo"
    package = ingest_cash_case(
        case_dir=case_dir,
        trial_balance_path=SAMPLE_DIR / "试算平衡表.xlsx",
        journal_path=SAMPLE_DIR / "银行存款日记账.xlsx",
        bank_statement_path=SAMPLE_DIR / "bank_confirmations",
        ocr_provider="pdf-text",
    )

    assert (case_dir / "case_package.json").exists()
    assert package.trial_balance
    assert package.bank_journal
    assert package.bank_confirmations
    assert package.trial_balance[0].source.document_hash


@pytest.mark.skipif(not SAMPLE_DIR.exists() or _external_template() is None, reason="external sample materials or reference C template not available")
def test_cash_workflow_writes_external_template() -> None:
    root = _runtime_dir()
    case_dir = root / "cash_demo"
    ingest_cash_case(
        case_dir=case_dir,
        trial_balance_path=SAMPLE_DIR / "试算平衡表.xlsx",
        journal_path=SAMPLE_DIR / "银行存款日记账.xlsx",
        bank_statement_path=SAMPLE_DIR / "bank_confirmations",
        ocr_provider="pdf-text",
    )
    output = root / "filled.xlsx"
    result = run_cash_workflow(case_dir, template_path=_external_template(), output_path=output)

    assert output.exists()
    assert Path(result.write_plan_path).exists()
    assert Path(result.findings_path).exists()
    assert Path(result.provenance_path or "").exists()


@pytest.mark.skipif(not SAMPLE_DIR.exists() or _external_template() is None, reason="external sample materials or reference C template not available")
def test_trace_events_cover_audit_pipeline() -> None:
    root = _runtime_dir()
    case_dir = root / "cash_demo"
    tracer = AuditTracer(enabled=False)
    ingest_cash_case(
        case_dir=case_dir,
        trial_balance_path=SAMPLE_DIR / "试算平衡表.xlsx",
        journal_path=SAMPLE_DIR / "银行存款日记账.xlsx",
        bank_statement_path=SAMPLE_DIR / "bank_confirmations",
        ocr_provider="pdf-text",
        tracer=tracer,
    )
    run_cash_workflow(case_dir, template_path=_external_template(), output_path=root / "filled.xlsx", tracer=tracer)

    stages = {event.stage for event in tracer.events}
    assert {"Sense", "OCR", "Logic", "Harness", "Output"}.issubset(stages)


def test_product_has_no_generators_directory() -> None:
    root = Path(__file__).resolve().parents[1]
    assert not (root / "auditpaper_agent" / "generators").exists()
