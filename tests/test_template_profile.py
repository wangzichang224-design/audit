from __future__ import annotations

from auditpaper_agent.knowledge.cash import detect_cash_profile


def test_reference_profile_resolves_confirmation_sheet_without_trailing_space() -> None:
    profile = detect_cash_profile(
        [
            "汇总",
            "C.00 Lead",
            "C.00 BKD",
            "C.01 Confirmations",
            "C.02 Bank reconciliations",
            "C.03 Cutoff",
        ]
    )

    assert profile.sheet_confirmations == "C.01 Confirmations"
    assert "C.01 Confirmations" in profile.allowed_cells
    assert "C.01 Confirmations " not in profile.allowed_cells
