from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class TraceEvent:
    stage: str
    message: str
    detail: str = ""


class AuditTracer:
    """Small trace adapter for CLI now and dashboard later."""

    def __init__(self, enabled: bool = False) -> None:
        self.enabled = enabled
        self.events: list[TraceEvent] = []
        self._console = None
        if enabled:
            try:
                from rich.console import Console

                self._console = Console()
            except Exception:
                self._console = None

    def emit(self, stage: str, message: str, detail: Any = "") -> None:
        text_detail = "" if detail is None else str(detail)
        event = TraceEvent(stage=stage, message=message, detail=text_detail)
        self.events.append(event)
        if not self.enabled:
            return
        prefix = f"[{stage}]"
        if self._console is not None:
            style = {
                "Sense": "cyan",
                "OCR": "magenta",
                "Logic": "yellow",
                "Harness": "green",
                "Output": "blue",
            }.get(stage, "white")
            line = f"[bold {style}]{prefix}[/bold {style}] {message}"
            if text_detail:
                line += f" [dim]{text_detail}[/dim]"
            self._console.print(line)
        else:
            print(f"{prefix} {message}" + (f" {text_detail}" if text_detail else ""))


def ensure_tracer(tracer: AuditTracer | None = None) -> AuditTracer:
    return tracer if tracer is not None else AuditTracer(enabled=False)
