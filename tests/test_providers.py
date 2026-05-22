from __future__ import annotations

from pathlib import Path

import pytest

from auditpaper_agent.config import ProviderConfig, Settings, TextinConfig
from auditpaper_agent.reasoning import DeepSeekReasoningProvider
from auditpaper_agent.sensing.ocr import TextinOcrProvider, QwenOcrProvider, normalize_confirmation_payload


def _settings_without_keys() -> Settings:
    return Settings(
        reasoning=ProviderConfig(provider="deepseek", base_url="https://example.invalid/v1", api_key="", model="deepseek-chat"),
        vision=ProviderConfig(provider="qwen", base_url="https://example.invalid/v1", api_key="", model="qwen-vl-max"),
        ocr_provider="pdf-text",
        textin=TextinConfig(app_id="", secret_code="", endpoint=""),
    )


def test_qwen_ocr_requires_api_key() -> None:
    provider = QwenOcrProvider(settings=_settings_without_keys())
    with pytest.raises(RuntimeError, match="API key"):
        provider.extract_confirmation(Path("工商银行_123456789012345_询证函回函.pdf"))


def test_textin_requires_credentials() -> None:
    provider = TextinOcrProvider(settings=_settings_without_keys())
    with pytest.raises(RuntimeError, match="Textin OCR requires"):
        provider.extract_confirmation(Path("工商银行_123456789012345_询证函回函.pdf"))


def test_deepseek_reasoning_requires_api_key() -> None:
    provider = DeepSeekReasoningProvider(settings=_settings_without_keys())
    with pytest.raises(RuntimeError, match="API key"):
        provider.enhance_findings(package=None, findings=[object()])  # type: ignore[arg-type,list-item]


def test_normalize_confirmation_payload_from_mock_response() -> None:
    payload = {
        "result": {
            "银行名称": "工商银行深圳分行",
            "银行账号": "5566 3344 2211 009",
            "币种": "CNY",
            "回函余额": "11,065,460.00",
            "受限金额": "0",
        }
    }

    normalized = normalize_confirmation_payload(payload, Path("工商银行深圳分行_556633442211009_询证函回函.pdf"))

    assert normalized["bank_name"] == "工商银行深圳分行"
    assert normalized["bank_account"] == "5566 3344 2211 009"
    assert normalized["confirmed_balance"] == 11065460.0
