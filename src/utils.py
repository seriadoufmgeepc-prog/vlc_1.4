from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional
from zoneinfo import ZoneInfo

BRAZIL_TZ = ZoneInfo("America/Sao_Paulo")
MONTHS = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
MONTH_INDEX = {m: i for i, m in enumerate(MONTHS)}
MONEY_RE = re.compile(r"\(?-?\d[\d\.,]*\)?")


LOWERCASE_WORDS = {
    "a", "ao", "aos", "à", "às", "ante", "após", "até", "com", "como", "contra", "da", "das", "de",
    "do", "dos", "e", "em", "entre", "na", "nas", "no", "nos", "ou", "para", "per", "pela", "pelas",
    "pelo", "pelos", "por", "sem", "sob", "sobre", "trás"
}


def current_brasilia_datetime() -> datetime:
    return datetime.now(BRAZIL_TZ)


def format_brasilia_datetime(value: datetime | None = None, with_seconds: bool = True) -> str:
    value = value or current_brasilia_datetime()
    if value.tzinfo is None:
        value = value.replace(tzinfo=BRAZIL_TZ)
    else:
        value = value.astimezone(BRAZIL_TZ)
    return value.strftime("%d/%m/%Y %H:%M:%S" if with_seconds else "%d/%m/%Y %H:%M")


def normalize_spaces(text: str) -> str:
    text = text.replace("\xa0", " ")
    return re.sub(r"[ \t]+", " ", text).strip()


def clean_money_token(token: str) -> str:
    token = token.strip().replace(" ", "")
    if token.startswith("(") and token.endswith(")"):
        token = "-" + token[1:-1]
    return token.replace(".", "").replace(",", ".")


def parse_money(token: str | int | float | None) -> float:
    if token is None:
        return 0.0
    if isinstance(token, (int, float)):
        return float(token)
    try:
        return float(clean_money_token(str(token)))
    except Exception:
        return 0.0


def format_br_number(value: float | int | None, decimals: int = 2) -> str:
    if value is None or value == "":
        return ""
    try:
        number = float(value)
    except Exception:
        return str(value)
    return f"{number:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def format_competence(month_index: int, year: int) -> str:
    return f"{MONTHS[month_index]}/{year}"


def competence_to_month_index(competence: Optional[str]) -> Optional[int]:
    if not competence:
        return None
    prefix = competence.strip().lower()[:3]
    return MONTH_INDEX.get(prefix)


def strip_numeric_noise_from_description(text: str) -> str:
    if not text:
        return ""
    cleaned = text.replace("*", " ")
    cleaned = re.sub(r"\(?\d{1,3}(?:\.\d{3})*,\d{2}\)?", " ", cleaned)
    cleaned = re.sub(r"\b\d+\b", " ", cleaned)
    cleaned = re.sub(r"\s{2,}", " ", cleaned)
    return cleaned.strip(" -")


def title_case_pt(text: str) -> str:
    text = normalize_spaces(text)
    if not text:
        return ""
    words = text.lower().split(" ")
    out: list[str] = []
    for idx, word in enumerate(words):
        if idx > 0 and word in LOWERCASE_WORDS:
            out.append(word)
        else:
            out.append(word[:1].upper() + word[1:])
    return " ".join(out)


def autosize_width(values: Iterable[object], minimum: int = 10, maximum: int = 42) -> float:
    size = minimum
    for value in values:
        size = max(size, len(str(value)) + 2)
    return min(size, maximum)


def file_signature(names: list[str]) -> str:
    return "|".join(sorted(names))


def safe_filename_part(text: str | None) -> str:
    text = (text or "arquivo").strip()
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)
    return text.strip("_") or "arquivo"


def file_suffix(path_or_name: str | Path) -> str:
    return Path(str(path_or_name)).suffix.lower()
