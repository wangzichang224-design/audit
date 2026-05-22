"""AuditPaper-Agent public SDK."""

from .logic.cash import run_cash_workflow
from .agent import check_provider_health, map_table_schema_with_agent
from .diagnostics import diagnose_client_package, run_workpaper_stress_suite
from .discovery import discover_cash_material_sets, resolve_cash_materials
from .sensing.cash import ingest_cash_case
from .sensing.erp import diagnose_erp_export, import_erp_export
from .service import inspect_cash_materials, run_auto_cash_case
from .suite import inspect_workpaper_project, run_auto_workpaper_suite, run_erp_workpaper_suite
from .workpaper_catalog import inspect_template_inventory

__all__ = [
    "diagnose_client_package",
    "diagnose_erp_export",
    "check_provider_health",
    "discover_cash_material_sets",
    "ingest_cash_case",
    "import_erp_export",
    "inspect_cash_materials",
    "inspect_template_inventory",
    "inspect_workpaper_project",
    "map_table_schema_with_agent",
    "resolve_cash_materials",
    "run_auto_cash_case",
    "run_auto_workpaper_suite",
    "run_cash_workflow",
    "run_erp_workpaper_suite",
    "run_workpaper_stress_suite",
]
