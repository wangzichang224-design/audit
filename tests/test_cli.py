from __future__ import annotations

from auditpaper_agent.cli import build_parser


def test_cli_parser_supports_cash_wizard() -> None:
    parser = build_parser()
    args = parser.parse_args(["wizard", "cash"])
    assert args.command == "wizard"
    assert args.case_type == "cash"


def test_cli_parser_supports_cash_auto() -> None:
    parser = build_parser()
    args = parser.parse_args(["auto", "cash", "--materials-dir", "case_folder"])
    assert args.command == "auto"
    assert args.case_type == "cash"
    assert args.materials_dir == "case_folder"


def test_cli_parser_supports_erp_diagnose_ingest_and_suite() -> None:
    parser = build_parser()
    diag = parser.parse_args(["diagnose", "erp", "--export-dir", "erp_folder", "--provider", "sap"])
    assert diag.command == "diagnose"
    assert diag.case_type == "erp"
    assert diag.provider == "sap"

    ingest = parser.parse_args(["ingest", "erp", "--export-dir", "erp_folder", "--case-dir", "case", "--confirm-mapping"])
    assert ingest.command == "ingest"
    assert ingest.case_type == "erp"
    assert ingest.confirm_mapping

    run = parser.parse_args(["run", "suite", "--erp-case-dir", "case"])
    assert run.command == "run"
    assert run.case_type == "suite"
    assert run.erp_case_dir == "case"
