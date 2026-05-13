from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd

from .utils import format_br_number

NUMERIC_COLUMNS = ["Saídas/Grupo", "Dep. Acumulada", "Vlr. Liq. Contábil"]
NOISE_RE = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}|\b\d{2,}\b")


@dataclass(frozen=True)
class QualityIssue:
    tipo: str
    mensagem: str


def _as_float(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def has_numeric_noise(text: object) -> bool:
    """Detecta provável contaminação de descrição por saldos/valores extraídos do PDF."""
    if text is None or pd.isna(text):
        return False
    return bool(NOISE_RE.search(str(text)))


def analyse_results(final_df: pd.DataFrame, memoria_df: pd.DataFrame | None = None) -> pd.DataFrame:
    """Gera uma análise sintética de consistência para exibição no app e exportação.

    A função não altera a apuração contábil. Ela apenas evidencia pontos de controle
    como duplicidades, ausência de mapeamento PCASP, campos textuais contaminados e
    coerência aritmética entre saídas, depreciação acumulada e valor líquido contábil.
    """
    issues: list[QualityIssue] = []

    if final_df.empty:
        issues.append(QualityIssue("Atenção", "Relatório Sintético sem grupos com saídas/baixas diferentes de zero."))
        return pd.DataFrame([issue.__dict__ for issue in issues])

    duplicated_groups = final_df.loc[final_df["Grupo"].duplicated(keep=False), "Grupo"].astype(str).unique().tolist()
    if duplicated_groups:
        issues.append(QualityIssue("Atenção", f"Há grupos duplicados no Relatório Sintético: {', '.join(duplicated_groups)}."))
    else:
        issues.append(QualityIssue("OK", "Não foram identificadas duplicidades de grupo no Relatório Sintético."))

    missing_accounts = final_df.loc[final_df["Conta Contábil"].fillna("").astype(str).str.strip().eq(""), "Grupo"].astype(str).tolist()
    if missing_accounts:
        issues.append(QualityIssue("Atenção", f"Grupos sem Conta Contábil mapeada: {', '.join(missing_accounts)}."))
    else:
        issues.append(QualityIssue("OK", "Todos os grupos do Relatório Sintético possuem Conta Contábil mapeada."))

    contaminated = final_df.loc[final_df["Descrição do Grupo"].map(has_numeric_noise), "Grupo"].astype(str).tolist()
    if contaminated:
        issues.append(QualityIssue("Atenção", f"Possível contaminação numérica na descrição dos grupos: {', '.join(contaminated)}."))
    else:
        issues.append(QualityIssue("OK", "Descrições dos grupos sem indícios de contaminação por valores numéricos."))

    diff = (_as_float(final_df["Saídas/Grupo"]) - _as_float(final_df["Dep. Acumulada"]) - _as_float(final_df["Vlr. Liq. Contábil"])).abs()
    max_diff = float(diff.max()) if not diff.empty else 0.0
    if max_diff > 0.01:
        issues.append(QualityIssue("Atenção", f"Diferença aritmética superior a R$ 0,01 encontrada na apuração do VLC. Maior diferença: {format_br_number(max_diff)}."))
    else:
        issues.append(QualityIssue("OK", "Apuração aritmética consistente: Saídas/Grupo - Dep. Acumulada = Vlr. Liq. Contábil."))

    for col in NUMERIC_COLUMNS:
        if col in final_df.columns:
            total = _as_float(final_df[col]).sum()
            issues.append(QualityIssue("Total", f"{col}: R$ {format_br_number(total)}."))

    if memoria_df is not None and not memoria_df.empty:
        groups_final = set(final_df["Grupo"].astype(str))
        groups_memory = set(memoria_df.loc[_as_float(memoria_df.get("Saídas/Grupo", pd.Series(dtype=float))) != 0, "Grupo"].astype(str))
        if groups_final == groups_memory:
            issues.append(QualityIssue("OK", "Grupos com saídas na Memória de Cálculo compatíveis com o Relatório Sintético."))
        else:
            only_final = sorted(groups_final - groups_memory)
            only_memory = sorted(groups_memory - groups_final)
            parts = []
            if only_final:
                parts.append(f"somente no Relatório Sintético: {', '.join(only_final)}")
            if only_memory:
                parts.append(f"somente na memória: {', '.join(only_memory)}")
            issues.append(QualityIssue("Atenção", "Diferença entre grupos do Relatório Sintético e da Memória de Cálculo: " + "; ".join(parts) + "."))

    return pd.DataFrame([issue.__dict__ for issue in issues])


def totals_dict(final_df: pd.DataFrame) -> dict[str, float]:
    return {col: float(_as_float(final_df[col]).sum()) if col in final_df.columns else 0.0 for col in NUMERIC_COLUMNS}
