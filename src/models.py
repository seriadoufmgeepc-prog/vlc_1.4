from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any

import pandas as pd


@dataclass
class ParsedDocument:
    doc_type: str
    filename: str
    ug_code: str | None = None
    ug_name: str | None = None
    period_label: str | None = None
    competence: str | None = None
    pages: int = 0
    rows: list[dict[str, Any]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_frame(self) -> pd.DataFrame:
        return pd.DataFrame(self.rows)

    def to_meta(self) -> dict[str, Any]:
        data = asdict(self)
        data.pop("rows", None)
        data.pop("warnings", None)
        return data
