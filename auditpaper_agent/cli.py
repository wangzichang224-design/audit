from __future__ import annotations

import argparse
from pathlib import Path

from auditpaper_agent.config import get_settings
from auditpaper_agent.discovery import discover_cash_materials
from auditpaper_agent.logic.cash import run_cash_workflow
from auditpaper_agent.sensing.cash import ingest_cash_case
from auditpaper_agent.sensing.erp import diagnose_erp_export, import_erp_export, write_erp_mapping_manifest
from auditpaper_agent.suite import run_erp_workpaper_suite
from auditpaper_agent.trace import AuditTracer
from auditpaper_agent.wizard import run_cash_wizard


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="auditpaper", description="AuditPaper-Agent CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    wizard = sub.add_parser("wizard", help="Interactive auditor-friendly workflow")
    wizard_sub = wizard.add_subparsers(dest="case_type", required=True)
    wizard_sub.add_parser("cash", help="Interactive cash workpaper wizard")

    auto = sub.add_parser("auto", help="One-folder automatic workflow")
    auto_sub = auto.add_subparsers(dest="case_type", required=True)
    cash_auto = auto_sub.add_parser("cash", help="Discover cash materials from one folder and run")
    cash_auto.add_argument("--materials-dir", required=True, help="Folder containing TB, journal, confirmations/statements, and template.")
    cash_auto.add_argument("--case-dir", default="")
    cash_auto.add_argument("--output", default="")
    cash_auto.add_argument("--ocr-provider", default="", choices=["", "pdf-text", "stub", "model-ocr", "qwen", "qwen-ocr", "textin"])
    cash_auto.add_argument("--trace", "--verbose", action="store_true", help="Show audit trace events.")
    cash_auto.add_argument("--use-reasoning", action="store_true", help="Use configured reasoning model to improve finding wording.")

    ingest = sub.add_parser("ingest", help="Ingest external client materials")
    ingest_sub = ingest.add_subparsers(dest="case_type", required=True)
    cash_ingest = ingest_sub.add_parser("cash", help="Ingest cash and bank materials")
    cash_ingest.add_argument("--case-dir", required=True)
    cash_ingest.add_argument("--trial-balance", required=True)
    cash_ingest.add_argument("--journal", required=True)
    cash_ingest.add_argument("--bank-statement", default="", help="Bank confirmation/statement file or directory")
    cash_ingest.add_argument("--ocr-provider", default="", choices=["", "pdf-text", "stub", "model-ocr", "qwen", "qwen-ocr", "textin"])
    cash_ingest.add_argument("--client-name", default="")
    cash_ingest.add_argument("--period-end", default="")
    cash_ingest.add_argument("--trace", "--verbose", action="store_true", help="Show audit trace events.")
    erp_ingest = ingest_sub.add_parser("erp", help="Ingest SAP/Yonyou ERP exports into a standardized audit package")
    erp_ingest.add_argument("--export-dir", required=True, help="Folder containing ERP export files.")
    erp_ingest.add_argument("--case-dir", required=True, help="Output folder for mapping_manifest.json and erp_package.json.")
    erp_ingest.add_argument("--provider", default="auto", choices=["auto", "sap", "yonyou"])
    erp_ingest.add_argument("--confirm-mapping", action="store_true", help="Confirm reviewed mapping_manifest.json and allow package creation.")

    diagnose = sub.add_parser("diagnose", help="Diagnose local client or ERP materials")
    diagnose_sub = diagnose.add_subparsers(dest="case_type", required=True)
    erp_diag = diagnose_sub.add_parser("erp", help="Build an ERP mapping manifest without importing")
    erp_diag.add_argument("--export-dir", required=True)
    erp_diag.add_argument("--output-dir", default="")
    erp_diag.add_argument("--provider", default="auto", choices=["auto", "sap", "yonyou"])

    run = sub.add_parser("run", help="Run audit workflow")
    run_sub = run.add_subparsers(dest="case_type", required=True)
    cash_run = run_sub.add_parser("cash", help="Run cash workflow and optionally fill Excel")
    cash_run.add_argument("--case-dir", required=True)
    cash_run.add_argument("--template", default="")
    cash_run.add_argument("--output", default="")
    cash_run.add_argument("--trace", "--verbose", action="store_true", help="Show audit trace events.")
    cash_run.add_argument("--use-reasoning", action="store_true", help="Use configured reasoning model to improve finding wording.")
    suite_run = run_sub.add_parser("suite", help="Generate a clean-room main-cycle workpaper suite")
    suite_run.add_argument("--erp-case-dir", required=True, help="Folder containing erp_package.json from `ingest erp`.")
    suite_run.add_argument("--output-dir", default="")

    validate = sub.add_parser("validate", help="Validate that expected case artifacts exist")
    validate.add_argument("--case-dir", required=True)
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "wizard" and args.case_type == "cash":
        run_cash_wizard()
        return

    if args.command == "auto" and args.case_type == "cash":
        discovery = discover_cash_materials(args.materials_dir)
        missing = discovery.missing_required
        if missing:
            raise SystemExit(f"资料文件夹缺少必要文件：{', '.join(missing)}")
        settings = get_settings()
        selected_ocr = args.ocr_provider or "pdf-text"
        case_dir = args.case_dir or str(Path("run_out") / f"{Path(args.materials_dir).name}_cash")
        output = args.output or str(Path(case_dir) / "filled_workpaper.xlsx")
        tracer = AuditTracer(enabled=args.trace)
        package = ingest_cash_case(
            case_dir=case_dir,
            trial_balance_path=discovery.trial_balance,
            journal_path=discovery.journal,
            bank_statement_path=discovery.bank_statement,
            ocr_provider=selected_ocr,
            tracer=tracer,
        )
        result = run_cash_workflow(
            case_dir=case_dir,
            template_path=discovery.template,
            output_path=output,
            tracer=tracer,
            use_reasoning=args.use_reasoning,
        )
        if not args.trace:
            print(f"Auto cash workflow completed: {args.materials_dir}")
            print(f"  client: {package.meta.client_name}")
            print(f"  findings: {len(result.findings)}")
            print(f"  output: {result.output_path}")
        return

    if args.command == "ingest" and args.case_type == "cash":
        tracer = AuditTracer(enabled=args.trace)
        settings = get_settings()
        selected_ocr = args.ocr_provider or settings.ocr_provider
        package = ingest_cash_case(
            case_dir=args.case_dir,
            trial_balance_path=args.trial_balance,
            journal_path=args.journal,
            bank_statement_path=args.bank_statement or None,
            ocr_provider=selected_ocr,
            client_name=args.client_name or None,
            period_end=args.period_end or None,
            tracer=tracer,
        )
        tracer.emit("Output", "ingest summary", f"tb={len(package.trial_balance)} journal={len(package.bank_journal)} confirmations={len(package.bank_confirmations)}")
        if not args.trace:
            print(f"Ingested cash case: {args.case_dir}")
            print(f"  trial_balance rows: {len(package.trial_balance)}")
            print(f"  bank_journal rows: {len(package.bank_journal)}")
            print(f"  confirmations rows: {len(package.bank_confirmations)}")
        return

    if args.command == "ingest" and args.case_type == "erp":
        result = import_erp_export(
            export_dir=args.export_dir,
            case_dir=args.case_dir,
            provider=args.provider,
            confirm_mapping=args.confirm_mapping,
        )
        if not result.success:
            print(f"ERP mapping manifest written: {result.manifest_path}")
            raise SystemExit("; ".join(result.errors))
        print(f"Ingested ERP export: {args.export_dir}")
        print(f"  mapping_manifest: {result.manifest_path}")
        print(f"  erp_package: {result.package_path}")
        if result.package:
            print(f"  tb rows: {len(result.package.trial_balance)}")
            print(f"  general ledger rows: {len(result.package.general_ledger)}")
        return

    if args.command == "diagnose" and args.case_type == "erp":
        output_dir = args.output_dir or str(Path("run_out") / f"{Path(args.export_dir).name}_erp_mapping")
        manifest_path = write_erp_mapping_manifest(
            export_dir=args.export_dir,
            output_dir=output_dir,
            provider=args.provider,
            confirmed=False,
        )
        manifest = diagnose_erp_export(args.export_dir, provider=args.provider, confirmed=False)
        print(f"ERP mapping manifest written: {manifest_path}")
        print(f"  tables: {len(manifest.tables)}")
        print(f"  blocking issues: {len(manifest.blocking_issues)}")
        if manifest.blocking_issues:
            print("  " + "; ".join(manifest.blocking_issues))
        print("  review the manifest, then run `auditpaper ingest erp --confirm-mapping ...`")
        return

    if args.command == "run" and args.case_type == "cash":
        if bool(args.template) != bool(args.output):
            parser.error("--template and --output must be provided together")
        tracer = AuditTracer(enabled=args.trace)
        result = run_cash_workflow(
            case_dir=args.case_dir,
            template_path=args.template or None,
            output_path=args.output or None,
            tracer=tracer,
            use_reasoning=args.use_reasoning,
        )
        tracer.emit("Output", "run summary", f"findings={len(result.findings)}")
        if not args.trace:
            print(f"Ran cash workflow: {args.case_dir}")
            print(f"  findings: {len(result.findings)}")
            print(f"  write_plan: {result.write_plan_path}")
            if result.output_path:
                print(f"  output: {result.output_path}")
        return

    if args.command == "run" and args.case_type == "suite":
        result = run_erp_workpaper_suite(
            erp_case_dir=args.erp_case_dir,
            output_dir=args.output_dir or None,
        )
        if not result.success:
            raise SystemExit("; ".join(result.errors) or "ERP suite generation failed")
        print(f"Generated ERP workpaper suite: {result.output_dir}")
        print(f"  workbooks: {len(result.workbooks)}")
        print(f"  zip: {result.zip_path}")
        return

    if args.command == "validate":
        case_dir = Path(args.case_dir)
        required = ["case_package.json", "audit_findings.json", "write_plan.json"]
        missing = [name for name in required if not (case_dir / name).exists()]
        if missing:
            raise SystemExit(f"Missing artifacts: {', '.join(missing)}")
        print(f"Valid case artifacts: {case_dir}")
        return

    parser.error("Unsupported command")


if __name__ == "__main__":
    main()
