from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol


@dataclass(frozen=True)
class CashTemplateProfile:
    name: str
    sheet_lead: str
    sheet_bkd: str
    sheet_confirmations: str
    sheet_recon: str
    sheet_cutoff: str
    allowed_cells: dict[str, list[str]]


def reference_cash_profile() -> CashTemplateProfile:
    """Reference C cash workpaper profile.

    The template itself is external. This profile only records the cells that
    AuditPaper-Agent is allowed to write during the first cash slice.
    """
    allowed: dict[str, list[str]] = {
        "C.00 Lead": [
            "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C15", "C16", "C17", "C18", "C19", "C32",
            "B38", "C38", "E38", "F38", "G38", "I38",
            "B39", "C39", "E39", "F39", "G39", "I39",
            "B40", "C40", "E40", "F40", "G40", "I40",
        ],
        "C.00 BKD": [],
        "C.01 Confirmations ": [],
        "C.02 Bank reconciliations": [
            "C17", "C18", "C19", "C20", "C21", "C24", "F24", "B10", "C10",
        ],
        "C.03 Cutoff": ["C8", "C10"],
    }
    for row in range(10, 16):
        allowed["C.00 BKD"].extend(f"{col}{row}" for col in "BCDEFGHIJ")
    for row in range(10, 18):
        allowed["C.01 Confirmations "].extend(f"{col}{row}" for col in "BCDEFGHIJKLMN")
    for row in range(26, 30):
        allowed["C.02 Bank reconciliations"].extend([f"B{row}", f"C{row}", f"D{row}", f"E{row}", f"F{row}", f"G{row}"])
    for row in range(31, 35):
        allowed["C.02 Bank reconciliations"].extend([f"B{row}", f"C{row}", f"D{row}", f"E{row}", f"F{row}", f"G{row}"])
    for row in range(20, 25):
        allowed["C.03 Cutoff"].extend(f"{col}{row}" for col in "BCDEFGHIJKLMN")
    for row in range(29, 34):
        allowed["C.03 Cutoff"].extend(f"{col}{row}" for col in "BCDEFGHIJKLMN")
    return CashTemplateProfile(
        name="reference_c_cash_v1",
        sheet_lead="C.00 Lead",
        sheet_bkd="C.00 BKD",
        sheet_confirmations="C.01 Confirmations ",
        sheet_recon="C.02 Bank reconciliations",
        sheet_cutoff="C.03 Cutoff",
        allowed_cells={sheet: sorted(set(cells)) for sheet, cells in allowed.items()},
    )


def neutral_cash_profile() -> CashTemplateProfile:
    allowed: dict[str, list[str]] = {
        "货币资金主表": [
            "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10", "B15", "B16", "B17", "B18", "B19", "B23",
            "A29", "B29", "C29", "D29", "E29", "F29", "H29", "J29",
            "A30", "B30", "C30", "D30", "E30", "F30", "H30", "J30",
            "A31", "B31", "C31", "D31", "E31", "F31", "H31", "J31",
        ],
        "货币资金明细": [],
        "银行余额调节": ["B5", "B6", "B10", "D10", "F10", "B11", "D11", "F11", "B12", "D12", "F12", "B19", "E19"],
        "截止性测试": ["B4", "B5", "B6", "B7"],
    }
    for row in range(12, 112):
        allowed["货币资金明细"].extend(f"{col}{row}" for col in "BCDEFGHIJKL")
    for row in list(range(21, 26)) + list(range(27, 32)):
        allowed["银行余额调节"].extend([f"A{row}", f"B{row}", f"C{row}", f"D{row}", f"E{row}", f"F{row}"])
    for row in list(range(13, 33)) + list(range(38, 58)):
        allowed["截止性测试"].extend(f"{col}{row}" for col in "ABCDEFGHIJKLMNO")
    return CashTemplateProfile(
        name="neutral_c_cash_v1",
        sheet_lead="货币资金主表",
        sheet_bkd="货币资金明细",
        sheet_confirmations="货币资金明细",
        sheet_recon="银行余额调节",
        sheet_cutoff="截止性测试",
        allowed_cells={sheet: sorted(set(cells)) for sheet, cells in allowed.items()},
    )


def simulated_cash_profile() -> CashTemplateProfile:
    allowed: dict[str, list[str]] = {
        "C.00 Lead": ["C2", "C3", "C4"],
        "C.00 BKD": ["C2", "C3", "C4"],
        "C.01 Confirmations": ["C2", "C3", "C4"],
        "C.02 Bank reconciliations": ["C2", "C3", "C4"],
        "C.03 Cutoff": ["C2", "C3", "C4"],
    }
    for row in range(9, 18):
        allowed["C.00 Lead"].extend(f"{col}{row}" for col in "BCDEFGHIJK")
    for row in range(19, 31):
        allowed["C.00 Lead"].extend(f"{col}{row}" for col in "BCDEFGHIJ")
    for row in range(9, 29):
        allowed["C.00 BKD"].extend(f"{col}{row}" for col in "BCDEFGHIJKL")
        allowed["C.01 Confirmations"].extend(f"{col}{row}" for col in "BCDEFGHIJKL")
        allowed["C.03 Cutoff"].extend(f"{col}{row}" for col in "BCDEFGHIJKL")
    for row in range(9, 25):
        allowed["C.02 Bank reconciliations"].extend(f"{col}{row}" for col in "BCDEFGHIJKL")
    for row in range(26, 36):
        allowed["C.02 Bank reconciliations"].extend(f"{col}{row}" for col in "BCDEFGHI")
    return CashTemplateProfile(
        name="sim_c_cash_v1",
        sheet_lead="C.00 Lead",
        sheet_bkd="C.00 BKD",
        sheet_confirmations="C.01 Confirmations",
        sheet_recon="C.02 Bank reconciliations",
        sheet_cutoff="C.03 Cutoff",
        allowed_cells={sheet: sorted(set(cells)) for sheet, cells in allowed.items()},
    )


class _WorkbookLike(Protocol):
    sheetnames: list[str]

    def __getitem__(self, key: str):
        ...


def detect_cash_profile_from_workbook(workbook: _WorkbookLike) -> CashTemplateProfile:
    sheet_names = list(workbook.sheetnames)
    if _looks_like_simulated_cash_template(workbook):
        return _adapt_profile_to_sheets(simulated_cash_profile(), sheet_names)
    return detect_cash_profile(sheet_names)


def detect_cash_profile(sheet_names: list[str]) -> CashTemplateProfile:
    if "C.00 Lead" in sheet_names and "C.02 Bank reconciliations" in sheet_names:
        return _adapt_profile_to_sheets(reference_cash_profile(), sheet_names)
    if "货币资金主表" in sheet_names and "银行余额调节" in sheet_names:
        return _adapt_profile_to_sheets(neutral_cash_profile(), sheet_names)
    raise ValueError("Unsupported cash template: expected reference C cash or neutral cash sheets")


def _adapt_profile_to_sheets(profile: CashTemplateProfile, sheet_names: list[str]) -> CashTemplateProfile:
    actual = {_normalize_sheet_name(name): name for name in sheet_names}

    def resolve(name: str) -> str:
        return actual.get(_normalize_sheet_name(name), name)

    replacements = {
        profile.sheet_lead: resolve(profile.sheet_lead),
        profile.sheet_bkd: resolve(profile.sheet_bkd),
        profile.sheet_confirmations: resolve(profile.sheet_confirmations),
        profile.sheet_recon: resolve(profile.sheet_recon),
        profile.sheet_cutoff: resolve(profile.sheet_cutoff),
    }
    allowed = {}
    for sheet, cells in profile.allowed_cells.items():
        resolved = replacements.get(sheet, resolve(sheet))
        if resolved in sheet_names:
            allowed[resolved] = cells
    return CashTemplateProfile(
        name=profile.name,
        sheet_lead=replacements[profile.sheet_lead],
        sheet_bkd=replacements[profile.sheet_bkd],
        sheet_confirmations=replacements[profile.sheet_confirmations],
        sheet_recon=replacements[profile.sheet_recon],
        sheet_cutoff=replacements[profile.sheet_cutoff],
        allowed_cells=allowed,
    )


def _normalize_sheet_name(name: str) -> str:
    return " ".join(str(name).strip().lower().split())


def _looks_like_simulated_cash_template(workbook: _WorkbookLike) -> bool:
    required = {"C.00 Lead", "C.00 BKD", "C.01 Confirmations", "C.02 Bank reconciliations", "C.03 Cutoff"}
    if not required.issubset(set(workbook.sheetnames)):
        return False
    lead = workbook["C.00 Lead"]
    bkd = workbook["C.00 BKD"]
    recon = workbook["C.02 Bank reconciliations"]
    lead_headers = {str(lead.cell(8, col).value or "") for col in range(2, 12)}
    bkd_headers = {str(bkd.cell(8, col).value or "") for col in range(2, 13)}
    recon_headers = {str(recon.cell(8, col).value or "") for col in range(2, 13)}
    return (
        {"TB期末数", "银行对账单余额", "调节后账面数"}.issubset(lead_headers)
        and {"账户ID", "期末账面余额", "对账单余额"}.issubset(bkd_headers)
        and {"银行对账单余额", "调节后账面余额", "账面余额"}.issubset(recon_headers)
    )


AUDIT_TICKS = {
    "TB": "已与科目余额表核对无误",
    "C01": "已与银行函证/回函资料核对",
    "Rx": "重新计算，无异常",
    "SL": "已与序时账核对",
}
