"""ПРАВКА #61/#63: OCR-тракт MinerU. Общие типы."""

from dataclasses import dataclass

SEVERITIES = ("critical", "warning", "info")


@dataclass(frozen=True)
class Finding:
    rule: str
    severity: str            # одно из SEVERITIES
    page: int | None         # 1-based; None — страница неизвестна
    snippet: str             # дословный фрагмент из markdown, по нему ищется место пометки
    suggestion: str | None = None
