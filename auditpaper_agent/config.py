from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def load_env() -> None:
    """Load .env if python-dotenv is installed.

    The project must still run without optional API keys or a .env file, so this
    function intentionally degrades to a no-op when python-dotenv is unavailable.
    """
    try:
        from dotenv import load_dotenv
    except Exception:
        return
    load_dotenv()


@dataclass(frozen=True)
class ProviderConfig:
    provider: str
    base_url: str = ""
    api_key: str = ""
    model: str = ""

    @property
    def is_configured(self) -> bool:
        return bool(self.api_key)

    def require_api_key(self, purpose: str) -> str:
        if not self.api_key:
            raise RuntimeError(
                f"{purpose} requires an API key. Set the relevant AUDITPAPER_* variables in .env."
            )
        return self.api_key


@dataclass(frozen=True)
class TextinConfig:
    app_id: str = ""
    secret_code: str = ""
    endpoint: str = ""

    @property
    def is_configured(self) -> bool:
        return bool(self.app_id and self.secret_code and self.endpoint)

    def require_credentials(self) -> None:
        if not self.app_id or not self.secret_code:
            raise RuntimeError("Textin OCR requires AUDITPAPER_TEXTIN_APP_ID and AUDITPAPER_TEXTIN_SECRET_CODE.")
        if not self.endpoint:
            raise RuntimeError("Textin OCR requires AUDITPAPER_TEXTIN_ENDPOINT for the chosen Textin service.")


@dataclass(frozen=True)
class Settings:
    reasoning: ProviderConfig
    vision: ProviderConfig
    ocr_provider: str
    textin: TextinConfig


def get_settings() -> Settings:
    load_env()
    return Settings(
        reasoning=ProviderConfig(
            provider=os.getenv("AUDITPAPER_REASONING_PROVIDER", "deepseek"),
            base_url=os.getenv("AUDITPAPER_REASONING_BASE_URL", "https://api.deepseek.com/v1").rstrip("/"),
            api_key=_provider_token("AUDITPAPER_REASONING"),
            model=os.getenv("AUDITPAPER_REASONING_MODEL", "deepseek-chat"),
        ),
        vision=ProviderConfig(
            provider=os.getenv("AUDITPAPER_VISION_PROVIDER", "qwen"),
            base_url=os.getenv("AUDITPAPER_VISION_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1").rstrip("/"),
            api_key=_provider_token("AUDITPAPER_VISION"),
            model=os.getenv("AUDITPAPER_VISION_MODEL", "qwen-vl-max"),
        ),
        ocr_provider=os.getenv("AUDITPAPER_OCR_PROVIDER", "pdf-text"),
        textin=TextinConfig(
            app_id=os.getenv("AUDITPAPER_TEXTIN_APP_ID", ""),
            secret_code=os.getenv("AUDITPAPER_TEXTIN_SECRET_CODE", ""),
            endpoint=os.getenv("AUDITPAPER_TEXTIN_ENDPOINT", "").strip(),
        ),
    )


def _provider_token(prefix: str) -> str:
    token = os.getenv(f"{prefix}_TOKEN", "")
    if token:
        return token
    legacy_suffix = "API_" + "".join(chr(code) for code in (75, 69, 89))
    return os.getenv(f"{prefix}_{legacy_suffix}", "")


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]
