from __future__ import annotations

import hashlib
import json
from datetime import date, datetime
from pathlib import Path
from typing import Any

from pydantic import BaseModel


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(to_jsonable(data), ensure_ascii=False, indent=2), encoding="utf-8")


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def to_jsonable(data: Any) -> Any:
    if isinstance(data, BaseModel):
        return data.model_dump(mode="json")
    if isinstance(data, (datetime, date)):
        return data.isoformat()
    if isinstance(data, Path):
        return str(data)
    if isinstance(data, list):
        return [to_jsonable(v) for v in data]
    if isinstance(data, tuple):
        return [to_jsonable(v) for v in data]
    if isinstance(data, dict):
        return {str(k): to_jsonable(v) for k, v in data.items()}
    return data


def parse_float(value: Any, default: float = 0.0) -> float:
    if value is None:
        return default
    text = str(value).strip().replace(",", "")
    if not text or text.lower() == "nan":
        return default
    try:
        return float(text)
    except ValueError:
        return default


def safe_filename(text: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in text)
    return cleaned[:80] or "case"
