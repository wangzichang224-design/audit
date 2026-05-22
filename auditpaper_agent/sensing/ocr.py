from __future__ import annotations

import base64
import json
import re
import uuid
import urllib.request
from abc import ABC, abstractmethod
from datetime import datetime
from pathlib import Path
from typing import Any

from auditpaper_agent.config import Settings, get_settings
from auditpaper_agent.trace import AuditTracer, ensure_tracer


class OcrProvider(ABC):
    """OCR/extraction provider abstraction for bank confirmations/statements."""

    name: str

    @abstractmethod
    def extract_confirmation(self, path: Path) -> dict[str, Any] | None:
        """Return confirmation fields or None if extraction fails."""


class StubOcrProvider(OcrProvider):
    name = "stub"

    def extract_confirmation(self, path: Path) -> dict[str, Any] | None:
        acct_match = re.search(r"_(\d{10,24})_", path.name)
        return {
            "bank_name": path.stem.split("_")[0],
            "bank_account": acct_match.group(1) if acct_match else "",
            "currency": "CNY",
            "confirmed_balance": 0.0,
            "confirmation_date": None,
            "provider_note": "stub provider; replace with model OCR for real extraction",
        }


class ModelOcrProvider(StubOcrProvider):
    name = "model-ocr"

    def extract_confirmation(self, path: Path) -> dict[str, Any] | None:
        return QwenOcrProvider().extract_confirmation(path)


class QwenOcrProvider(OcrProvider):
    name = "qwen-ocr"

    def __init__(self, settings: Settings | None = None, tracer: AuditTracer | None = None) -> None:
        self.settings = settings or get_settings()
        self.tracer = ensure_tracer(tracer)

    def extract_confirmation(self, path: Path) -> dict[str, Any] | None:
        cfg = self.settings.vision
        cfg.require_api_key("Qwen OCR")
        self.tracer.emit("OCR", "calling Qwen vision OCR", f"model={cfg.model} file={path.name}")
        content = _openai_compatible_vision_call(
            base_url=cfg.base_url,
            api_key=cfg.api_key,
            model=cfg.model,
            path=path,
            prompt=_CONFIRMATION_JSON_PROMPT,
        )
        return normalize_confirmation_payload(extract_json_object(content), path)


class TextinOcrProvider(OcrProvider):
    name = "textin"

    def __init__(self, settings: Settings | None = None, tracer: AuditTracer | None = None) -> None:
        self.settings = settings or get_settings()
        self.tracer = ensure_tracer(tracer)

    def extract_confirmation(self, path: Path) -> dict[str, Any] | None:
        cfg = self.settings.textin
        cfg.require_credentials()
        self.tracer.emit("OCR", "calling Textin OCR", f"file={path.name}")
        data = _post_multipart(
            cfg.endpoint,
            fields={"config": json.dumps({"provider": "textin", "parse_mode": "auto"}, ensure_ascii=False)},
            files={"file": (path.name, path.read_bytes(), _guess_content_type(path))},
            headers={
                "x-ti-app-id": cfg.app_id,
                "x-ti-secret-code": cfg.secret_code,
            },
        )
        return normalize_confirmation_payload(data, path)


class PdfTextProvider(OcrProvider):
    name = "pdf-text"
    _amount_re = re.compile(r"\d{1,3}(?:,\d{3})+\.\d{2}|\b\d+\.\d{2}\b")
    _account_re = re.compile(r"(\d{4})\s*(\d{4})\s*(\d{4})\s*(\d{3,6})")
    _date_re = re.compile(r"(\d{4}[-/]\d{2}[-/]\d{2})")

    def extract_confirmation(self, path: Path) -> dict[str, Any] | None:
        try:
            from pypdf import PdfReader
        except Exception:
            return None
        try:
            reader = PdfReader(str(path))
            text = "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception:
            return None

        result: dict[str, Any] = {}
        name_match = re.match(r"(.+?)_\d+_", path.name)
        result["bank_name"] = name_match.group(1).replace("_", "") if name_match else path.stem

        acct_match = self._account_re.search(text)
        if acct_match:
            result["bank_account"] = "".join(acct_match.groups())
        else:
            fn_match = re.search(r"_(\d{10,24})_", path.name)
            result["bank_account"] = fn_match.group(1) if fn_match else ""

        amounts = [float(n.replace(",", "")) for n in self._amount_re.findall(text)]
        result["confirmed_balance"] = max(amounts) if amounts else 0.0
        result["currency"] = "USD" if ("USD" in text or "$" in text) else "CNY"

        dates = self._date_re.findall(text)
        if dates:
            for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
                try:
                    result["confirmation_date"] = datetime.strptime(dates[0], fmt).date()
                    break
                except ValueError:
                    pass
        return result if result.get("bank_account") or result.get("confirmed_balance") else None


_CONFIRMATION_JSON_PROMPT = (
    "Extract a Chinese bank confirmation reply into strict JSON only. "
    "Return keys: bank_name, bank_account, currency, confirmed_balance, "
    "confirmation_date, restricted_amount, restriction_nature. "
    "Use numbers for amounts. If a field is absent, use an empty string or 0."
)


def get_ocr_provider(name: str, settings: Settings | None = None, tracer: AuditTracer | None = None) -> OcrProvider:
    normalized = (name or "pdf-text").strip().lower()
    if normalized in {"pdf-text", "pdf"}:
        return PdfTextProvider()
    if normalized in {"stub", "none"}:
        return StubOcrProvider()
    if normalized in {"model-ocr", "qwen", "qwen-ocr"}:
        return QwenOcrProvider(settings=settings, tracer=tracer)
    if normalized in {"textin", "textin-ocr"}:
        return TextinOcrProvider(settings=settings, tracer=tracer)
    raise ValueError(f"Unknown OCR provider: {name}")


def normalize_confirmation_payload(data: dict[str, Any], path: Path | None = None) -> dict[str, Any]:
    """Normalize direct, Qwen, or Textin-shaped OCR payloads into confirmation fields."""
    raw: Any = data
    if "choices" in data:
        raw = extract_json_object(str(data["choices"][0]["message"]["content"]))
    elif "result" in data:
        raw = data["result"]
    elif "data" in data:
        raw = data["data"]
    if isinstance(raw, str):
        raw = extract_json_object(raw)
    if isinstance(raw, dict) and "structured" in raw:
        raw = raw["structured"]
    text = _text_from_textin_payload(data)
    if not isinstance(raw, dict):
        raw = {}

    file_hint = path.stem if path else ""
    acct_match = re.search(r"_(\d{10,24})_", path.name) if path else None
    bank_name = _first_value(raw, "bank_name", "银行名称", "bank", default="") or _infer_bank_name(text, file_hint)
    bank_account = _first_value(raw, "bank_account", "银行账号", "账号", "account_no", default="") or _infer_account(text)
    return {
        "bank_name": str(bank_name or file_hint.split("_")[0]),
        "bank_account": str(bank_account or (acct_match.group(1) if acct_match else "")),
        "currency": str(_first_value(raw, "currency", "币种", default="CNY") or "CNY"),
        "confirmed_balance": _parse_amount(_first_value(raw, "confirmed_balance", "回函余额", "确认余额", "银行余额", default=0.0)) or _infer_largest_amount(text),
        "confirmation_date": _first_value(raw, "confirmation_date", "回函日期", "函证日期", default=None) or _infer_date(text),
        "restricted_amount": _parse_amount(_first_value(raw, "restricted_amount", "受限金额", default=0.0)),
        "restriction_nature": str(_first_value(raw, "restriction_nature", "受限原因", "受限性质", default="") or _infer_restriction(text)),
    }


def extract_json_object(content: str) -> dict[str, Any]:
    text = content.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text, flags=re.IGNORECASE).strip()
        text = re.sub(r"```$", "", text).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, flags=re.DOTALL)
        if not match:
            raise ValueError(f"OCR provider did not return JSON: {content[:200]!r}")
        return json.loads(match.group(0))


def _first_value(data: dict[str, Any], *keys: str, default: Any = None) -> Any:
    for key in keys:
        if key in data and data[key] not in (None, ""):
            return data[key]
    return default


def _parse_amount(value: Any) -> float:
    if value is None:
        return 0.0
    text = str(value).replace(",", "").strip()
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        nums = re.findall(r"-?\d+(?:\.\d+)?", text)
        return float(nums[0]) if nums else 0.0


def _text_from_textin_payload(data: dict[str, Any]) -> str:
    def collect(value: Any) -> list[str]:
        if isinstance(value, dict):
            parts: list[str] = []
            if isinstance(value.get("text"), str):
                parts.append(value["text"])
            if isinstance(value.get("markdown"), str):
                parts.append(value["markdown"])
            for key in ("elements", "pages", "tables", "cells", "children"):
                if key in value:
                    parts.extend(collect(value[key]))
            return parts
        if isinstance(value, list):
            parts = []
            for item in value:
                parts.extend(collect(item))
            return parts
        return []

    return "\n".join(collect(data))


def _infer_bank_name(text: str, file_hint: str) -> str:
    match = re.search(r"([\u4e00-\u9fa5A-Za-z]+银行[\u4e00-\u9fa5A-Za-z]*)", text)
    if match:
        return match.group(1)
    return file_hint.split("_")[0] if file_hint else ""


def _infer_account(text: str) -> str:
    match = re.search(r"(\d{4})\s*(\d{4})\s*(\d{4})\s*(\d{3,6})", text)
    return "".join(match.groups()) if match else ""


def _infer_largest_amount(text: str) -> float:
    values = [_parse_amount(x) for x in re.findall(r"\d{1,3}(?:,\d{3})+\.\d{2}|\b\d+\.\d{2}\b", text)]
    return max(values) if values else 0.0


def _infer_date(text: str) -> str | None:
    match = re.search(r"\d{4}[-/年]\d{1,2}[-/月]\d{1,2}", text)
    return match.group(0).replace("年", "-").replace("月", "-") if match else None


def _infer_restriction(text: str) -> str:
    for keyword in ("质押", "冻结", "保证金", "受限", "监管"):
        if keyword in text:
            return keyword
    return ""


def _openai_compatible_vision_call(base_url: str, api_key: str, model: str, path: Path, prompt: str) -> str:
    mime = "application/pdf" if path.suffix.lower() == ".pdf" else "image/png"
    data_url = f"data:{mime};base64,{base64.b64encode(path.read_bytes()).decode('ascii')}"
    payload = {
        "model": model,
        "temperature": 0,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": data_url}},
                ],
            }
        ],
    }
    response = _post_json(
        f"{base_url.rstrip('/')}/chat/completions",
        payload,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
    )
    try:
        return str(response["choices"][0]["message"]["content"])
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError(f"Unexpected OpenAI-compatible OCR response: {response}") from exc


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
        raise RuntimeError(f"OCR HTTP request failed for {url}: {exc}") from exc


def _post_multipart(
    url: str,
    fields: dict[str, str],
    files: dict[str, tuple[str, bytes, str]],
    headers: dict[str, str],
) -> dict[str, Any]:
    boundary = f"----AuditPaperAgent{uuid.uuid4().hex}"
    body = bytearray()
    for name, value in fields.items():
        body.extend(f"--{boundary}\r\n".encode("utf-8"))
        body.extend(f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode("utf-8"))
        body.extend(value.encode("utf-8"))
        body.extend(b"\r\n")
    for name, (filename, content, content_type) in files.items():
        body.extend(f"--{boundary}\r\n".encode("utf-8"))
        body.extend(
            f'Content-Disposition: form-data; name="{name}"; filename="{filename}"\r\n'.encode("utf-8")
        )
        body.extend(f"Content-Type: {content_type}\r\n\r\n".encode("utf-8"))
        body.extend(content)
        body.extend(b"\r\n")
    body.extend(f"--{boundary}--\r\n".encode("utf-8"))
    request_headers = dict(headers)
    request_headers["Content-Type"] = f"multipart/form-data; boundary={boundary}"
    request = urllib.request.Request(url, data=bytes(body), headers=request_headers, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        raise RuntimeError(f"Textin multipart OCR request failed for {url}: {exc}") from exc
    if payload.get("code") not in (None, 200):
        raise RuntimeError(f"Textin OCR returned code={payload.get('code')} msg={payload.get('msg')}")
    return payload


def _guess_content_type(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return "application/pdf"
    if suffix in {".jpg", ".jpeg"}:
        return "image/jpeg"
    if suffix == ".png":
        return "image/png"
    return "application/octet-stream"
