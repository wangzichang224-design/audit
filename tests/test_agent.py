from __future__ import annotations

from auditpaper_agent.agent import check_provider_health
from auditpaper_agent.config import ProviderConfig, Settings, TextinConfig


def test_provider_health_reports_missing_keys_without_live_call() -> None:
    settings = Settings(
        reasoning=ProviderConfig(provider="deepseek", base_url="https://example.invalid/v1", api_key="", model="deepseek-chat"),
        vision=ProviderConfig(provider="qwen", base_url="https://example.invalid/v1", api_key="", model="qwen-vl-max"),
        ocr_provider="pdf-text",
        textin=TextinConfig(),
    )

    health = check_provider_health(live=False, settings=settings)

    assert not health.checked_live
    assert not health.reasoning.configured
    assert not health.reasoning.ok
    assert not health.vision.configured
    assert not health.vision.ok


def test_provider_health_reports_configured_keys_without_printing_secret() -> None:
    settings = Settings(
        reasoning=ProviderConfig(provider="deepseek", base_url="https://example.invalid/v1", api_key="secret", model="deepseek-chat"),
        vision=ProviderConfig(provider="qwen", base_url="https://example.invalid/v1", api_key="secret", model="qwen-vl-max"),
        ocr_provider="pdf-text",
        textin=TextinConfig(),
    )

    health = check_provider_health(live=False, settings=settings)

    assert health.reasoning.configured
    assert health.reasoning.ok
    assert health.vision.configured
    assert health.vision.ok
    assert "secret" not in repr(health)
