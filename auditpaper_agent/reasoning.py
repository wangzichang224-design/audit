from __future__ import annotations

import json
import re
import urllib.request
from typing import Any

from auditpaper_agent.config import Settings, get_settings
from auditpaper_agent.contracts import AuditFinding, StandardizedAuditPackage
from auditpaper_agent.trace import AuditTracer, ensure_tracer


class DeepSeekReasoningProvider:
    """OpenAI-compatible reasoning provider for audit wording only."""

    name = "deepseek"

    def __init__(self, settings: Settings | None = None, tracer: AuditTracer | None = None) -> None:
        self.settings = settings or get_settings()
        self.tracer = ensure_tracer(tracer)

    def enhance_findings(self, package: StandardizedAuditPackage, findings: list[AuditFinding]) -> list[AuditFinding]:
        cfg = self.settings.reasoning
        cfg.require_api_key("DeepSeek reasoning")
        if not findings:
            return findings
        self.tracer.emit("Logic", "calling DeepSeek reasoning provider", f"model={cfg.model} findings={len(findings)}")
        payload = {
            "model": cfg.model,
            "temperature": 0.1,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "You are a Chinese CPA audit senior. Improve only the wording of audit findings. "
                        "Do not change finding_type, severity, amount, or sources. Return strict JSON array."
                    ),
                },
                {
                    "role": "user",
                    "content": json.dumps(
                        {
                            "client_name": package.meta.client_name,
                            "period_end": package.meta.period_end.isoformat() if package.meta.period_end else "",
                            "findings": [f.model_dump(mode="json") for f in findings],
                        },
                        ensure_ascii=False,
                    ),
                },
            ],
        }
        response = _post_json(
            f"{cfg.base_url.rstrip('/')}/chat/completions",
            payload,
            headers={"Authorization": f"Bearer {cfg.api_key}", "Content-Type": "application/json"},
        )
        content = response["choices"][0]["message"]["content"]
        data = _extract_json_array(str(content))
        enhanced: list[AuditFinding] = []
        for idx, original in enumerate(findings):
            item = data[idx] if idx < len(data) and isinstance(data[idx], dict) else {}
            enhanced.append(original.model_copy(update={"description": str(item.get("description") or original.description)}))
        return enhanced


def _post_json(url: str, payload: dict[str, Any], headers: dict[str, str]) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            return json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        raise RuntimeError(f"Reasoning HTTP request failed for {url}: {exc}") from exc


def _extract_json_array(content: str) -> list[Any]:
    text = content.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text, flags=re.IGNORECASE).strip()
        text = re.sub(r"```$", "", text).strip()
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        start = text.find("[")
        end = text.rfind("]")
        if start < 0 or end < start:
            raise ValueError(f"Reasoning provider did not return a JSON array: {content[:200]!r}")
        data = json.loads(text[start : end + 1])
    if not isinstance(data, list):
        raise ValueError("Reasoning provider response must be a JSON array.")
    return data
