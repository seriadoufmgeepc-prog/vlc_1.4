from __future__ import annotations

import io
import re
from typing import BinaryIO

import pdfplumber

from .models import ParsedDocument
from .utils import MONTHS, competence_to_month_index, format_br_number, format_competence, normalize_spaces, parse_money, strip_numeric_noise_from_description

RSP_ROW_RE = re.compile(r"^\s*(\d{1,3})\s+(.+)$")
MONEY_VALUE_RE = re.compile(r"\(?-?\d{1,3}(?:\.\d{3})*,\d{2}\)?")
GROUP_HEADER_RE = re.compile(r"^\s*(\d{1,3})\s*-\s*(.+?)\s*$")
UG_RE = re.compile(r"(\d{6})\s*-\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ /.-]+)")
PERIOD_RE = re.compile(r"DE\s+(\d{1,2}/\s*\d{2}/\d{4})\s+A\s+(\d{1,2}/\s*\d{2}\s*/\d{4})", re.I)


def _read_pdf_pages(file_obj: BinaryIO) -> list[str]:
    data = file_obj.read()
    file_obj.seek(0)
    pages: list[str] = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = [normalize_spaces(line) for line in text.splitlines() if normalize_spaces(line)]
            pages.append("\n".join(lines))
    return pages


def detect_document_type(pages: list[str]) -> str:
    joined = "\n".join(pages).upper()
    if "RELATÓRIO SINTÉTICO PATRIMONIAL" in joined or "RELATORIO SINTETICO PATRIMONIAL" in joined:
        return "rsp"
    if "RELATÓRIO DE DEPRECIAÇÃO ACUMULADA" in joined or "RELATORIO DE DEPRECIACAO ACUMULADA" in joined:
        return "depreciacao"
    return "desconhecido"


def extract_metadata(pages: list[str]) -> tuple[str | None, str | None, str | None, str | None]:
    joined = "\n".join(pages)
    ug_code = None
    ug_name = None
    period_label = None
    competence = None

    match_ug = UG_RE.search(joined)
    if match_ug:
        ug_code = match_ug.group(1)
        ug_name = normalize_spaces(match_ug.group(2).strip())

    match_period = PERIOD_RE.search(joined)
    if match_period:
        d1 = re.sub(r"\s+", "", match_period.group(1))
        d2 = re.sub(r"\s+", "", match_period.group(2))
        period_label = f"{d1} a {d2}"
        month = int(d1.split("/")[1])
        year = int(d1.split("/")[2])
        competence = format_competence(month - 1, year)

    if not competence:
        year_match = re.search(r"EXERC[ÍI]CIO\s+(20\d{2})", joined, re.I)
        year = int(year_match.group(1)) if year_match else None
        if year:
            upper = joined.upper()
            for idx, abbr in enumerate([m.upper() for m in MONTHS]):
                if re.search(rf"\b{abbr}\b", upper):
                    competence = format_competence(idx, year)
                    break
    return ug_code, ug_name, period_label, competence


def _parse_rsp_saida_detalhada(pages: list[str]) -> list[dict]:
    """Extrai o quadro "Relatório Sintético Patrimonial de Saídas".

    Regra de negócio v10: a apuração deve evidenciar somente as demais saídas
    patrimoniais, desconsiderando saídas por transferência, pois o SICPAT já
    possui relatório próprio para transferência com valor do bem e depreciação
    acumulada associada. No layout do quadro de saídas, o valor sem
    transferência corresponde à coluna "SAÍDAS"; o "TOTAL SAÍDAS" inclui
    a coluna "TRANSFERE SAÍDAS".
    """
    rows: list[dict] = []
    for page in pages:
        upper_page = page.upper()
        if not (
            "RELATÓRIO SINTÉTICO PATRIMONIAL DE SAÍDAS" in upper_page
            or "RELATORIO SINTETICO PATRIMONIAL DE SAIDAS" in upper_page
        ):
            continue
        for raw_line in page.splitlines():
            line = normalize_spaces(raw_line)
            match = RSP_ROW_RE.match(line)
            if not match:
                continue
            grupo = int(match.group(1))
            remainder = match.group(2)
            money_matches = list(MONEY_VALUE_RE.finditer(remainder))
            if len(money_matches) < 9:
                continue
            first_money = money_matches[0]
            descricao = strip_numeric_noise_from_description(remainder[: first_money.start()]) or f"Grupo {grupo}"
            nums = [parse_money(m.group(0)) for m in money_matches]
            # Layout do quadro de saídas extraído do SICPAT:
            # alienação, doação, abandono, lançamento indevido, outros,
            # SAÍDAS (subtotal sem transferências), transfere comodato,
            # transfere saídas, TOTAL SAÍDAS, comodato.
            doacao = nums[1] if len(nums) > 1 else 0.0
            saidas_sem_transferencia = nums[5]
            transferencias_saida = nums[7] if len(nums) > 7 else 0.0
            total_saidas_original = nums[8] if len(nums) > 8 else saidas_sem_transferencia + transferencias_saida
            rows.append({
                "Grupo": grupo,
                "Descrição do Grupo": descricao,
                "Saídas/Grupo": saidas_sem_transferencia,
                "Doação": doacao,
                "Saídas por Transferência": transferencias_saida,
                "Total Saídas Original": total_saidas_original,
                "Saldo Anterior": 0.0,
                "Entradas": 0.0,
                "Reavaliação": 0.0,
                "Red.Vlr.Rec.": 0.0,
                "Saldo Atual": 0.0,
                "Depreciação do RSP": 0.0,
                "Vlr. Liq. Cont. do RSP": 0.0,
            })
    return rows


def _parse_rsp_sintetico_geral(pages: list[str]) -> list[dict]:
    """Fallback para PDFs antigos sem quadro detalhado de saídas.

    Neste caso, não é possível separar transferências com segurança; por isso,
    a coluna de saídas do sintético geral é usada apenas como contingência.
    """
    rows: list[dict] = []
    for page in pages:
        upper_page = page.upper()
        if "RELATÓRIO SINTÉTICO PATRIMONIAL DE ENTRADAS" in upper_page or "RELATORIO SINTETICO PATRIMONIAL DE ENTRADAS" in upper_page:
            continue
        if "RELATÓRIO SINTÉTICO PATRIMONIAL DE SAÍDAS" in upper_page or "RELATORIO SINTETICO PATRIMONIAL DE SAIDAS" in upper_page:
            continue
        for raw_line in page.splitlines():
            line = normalize_spaces(raw_line)
            match = RSP_ROW_RE.match(line)
            if not match:
                continue
            grupo = int(match.group(1))
            remainder = match.group(2)
            money_matches = list(MONEY_VALUE_RE.finditer(remainder))
            if len(money_matches) < 8:
                continue
            first_money = money_matches[0]
            descricao = strip_numeric_noise_from_description(remainder[: first_money.start()]) or f"Grupo {grupo}"
            nums = [parse_money(m.group(0)) for m in money_matches]
            rows.append({
                "Grupo": grupo,
                "Descrição do Grupo": descricao,
                "Saldo Anterior": nums[0],
                "Entradas": nums[1],
                "Saídas/Grupo": nums[2],
                "Doação": 0.0,
                "Saídas por Transferência": 0.0,
                "Total Saídas Original": nums[2],
                "Reavaliação": nums[3],
                "Red.Vlr.Rec.": nums[4],
                "Saldo Atual": nums[5],
                "Depreciação do RSP": nums[6],
                "Vlr. Liq. Cont. do RSP": nums[7],
            })
    return rows


def parse_rsp(file_obj: BinaryIO, filename: str) -> ParsedDocument:
    pages = _read_pdf_pages(file_obj)
    ug_code, ug_name, period_label, competence = extract_metadata(pages)
    warnings: list[str] = []

    rows = _parse_rsp_saida_detalhada(pages)
    if rows:
        total_original = sum(float(row.get("Total Saídas Original", 0.0)) for row in rows)
        total_considerado = sum(float(row.get("Saídas/Grupo", 0.0)) for row in rows)
        total_transferencias = sum(float(row.get("Saídas por Transferência", 0.0)) for row in rows)
        total_doacoes = sum(float(row.get("Doação", 0.0)) for row in rows)
        if abs(total_doacoes) > 0.01:
            warnings.append(
                "Saídas por doação foram identificadas e serão segregadas na aba Registro SIAFI Web: "
                f"doações=R$ {format_br_number(total_doacoes)}."
            )
        if abs(total_transferencias) > 0.01:
            warnings.append(
                "Saídas por transferência foram desconsideradas da apuração: "
                f"total original=R$ {format_br_number(total_original)}; transferências=R$ {format_br_number(total_transferencias)}; "
                f"base considerada=R$ {format_br_number(total_considerado)}."
            )
    else:
        rows = _parse_rsp_sintetico_geral(pages)
        if rows:
            warnings.append(
                "O quadro detalhado de saídas não foi localizado. O sistema utilizou o quadro sintético geral como contingência; "
                "nesse modo não é possível separar saídas por transferência ou doação com segurança."
            )

    if not rows:
        warnings.append("Nenhuma linha do Relatório Sintético Patrimonial foi identificada pelo parser.")
    return ParsedDocument("rsp", filename, ug_code, ug_name, period_label, competence, len(pages), rows, warnings)


def _has_depreciation_table_data(block_lines: list[str]) -> bool:
    """Indica se o bloco contém linhas numéricas reais da tabela mensal.

    Cabeçalhos repetidos no rodapé de uma página podem ser extraídos sem os
    valores da tabela e continuar na página seguinte. Nesses casos, não devemos
    gerar uma linha parcial nem usar números de página, datas ou exercício como
    valores monetários.
    """
    return any(len(MONEY_VALUE_RE.findall(line)) >= 10 for line in block_lines)


def _is_page_or_report_header(line: str) -> bool:
    """Identifica linhas de cabeçalho/rodapé que nunca devem ser lidas como valores.

    A correção v12 remove qualquer fallback que pudesse interpretar o ano
    do cabeçalho (ex.: EXERCÍCIO 2026) como valor monetário.
    """
    upper = line.upper()
    header_markers = [
        "RELATÓRIO DE DEPRECIAÇÃO ACUMULADA",
        "RELATORIO DE DEPRECIACAO ACUMULADA",
        "UNIVERSIDADE FEDERAL DE MINAS GERAIS",
        "SICPAT - SISTEMA DE CONTROLE PATRIMONIAL",
        "HOSPITAL DAS CLÍNICAS",
        "HOSPITAL DAS CLINICAS",
        "PÁGINA",
        "PAGINA",
        "EXERCÍCIO",
        "EXERCICIO",
    ]
    return any(marker in upper for marker in header_markers)


def _block_to_saida_baixas(block_lines: list[str], month_index: int) -> tuple[float, bool]:
    """Retorna a linha SAÍDAS (BAIXAS) do bloco do grupo.

    Regra v12: somente a linha explicitamente rotulada como SAÍDAS (BAIXAS)
    é fonte válida para a Dep. Acumulada (IMB010). Não há fallback por posição
    de linha numérica, pois em quebras de página o cabeçalho pode trazer o ano
    2026 e contaminar a leitura.
    """
    for line in block_lines:
        if _is_page_or_report_header(line):
            continue
        upper = line.upper()
        if "SAÍDAS (BAIXAS)" in upper or "SAIDAS (BAIXAS)" in upper:
            tokens = MONEY_VALUE_RE.findall(line)
            # A linha mensal deve conter os 12 meses. Aceitar menos que isso
            # aumenta o risco de capturar cabeçalhos, datas ou fragmentos.
            if len(tokens) >= 12 and month_index < len(tokens):
                vals = [parse_money(tok) for tok in tokens]
                return abs(vals[month_index]), True
            return 0.0, False
    return 0.0, False


def parse_depreciacao(file_obj: BinaryIO, filename: str, competence_hint: str | None = None) -> ParsedDocument:
    pages = _read_pdf_pages(file_obj)
    ug_code, ug_name, period_label, competence = extract_metadata(pages)
    competence = competence_hint or competence
    month_index = competence_to_month_index(competence)
    if month_index is None:
        month_index = 0

    rows: list[dict] = []
    warnings: list[str] = []
    current_group: int | None = None
    current_desc = ""
    block_lines: list[str] = []

    def flush_block() -> None:
        nonlocal current_group, current_desc, block_lines
        if current_group is None:
            return
        value, found = _block_to_saida_baixas(block_lines, month_index)
        if not found and not _has_depreciation_table_data(block_lines):
            # Cabeçalho repetido no final de página, sem dados efetivos.
            current_group = None
            current_desc = ""
            block_lines = []
            return
        if not found:
            warnings.append(f"Grupo {current_group}: linha de SAÍDAS (BAIXAS) não encontrada com segurança; valor assumido como zero.")
        rows.append({
            "Grupo": current_group,
            "Descrição do Grupo": strip_numeric_noise_from_description(current_desc) or f"Grupo {current_group}",
            "Dep. Acumulada (IMB010)": value,
        })
        current_group = None
        current_desc = ""
        block_lines = []

    for page in pages:
        for raw_line in page.splitlines():
            line = normalize_spaces(raw_line)
            header = GROUP_HEADER_RE.match(line)
            if header:
                flush_block()
                current_group = int(header.group(1))
                current_desc = header.group(2)
                block_lines = []
                continue
            if current_group is not None:
                block_lines.append(line)
    flush_block()

    # Segurança adicional: em PDFs com quebra de página entre o cabeçalho do
    # grupo e sua tabela, manter apenas uma ocorrência por grupo. Como a regra
    # acima só aceita SAÍDAS (BAIXAS) rotulada, duplicidades remanescentes são
    # tratadas sem soma, preservando o valor mais recente do grupo.
    if rows:
        dedup: dict[int, dict] = {}
        duplicate_groups: set[int] = set()
        for row in rows:
            grupo = int(row["Grupo"])
            if grupo in dedup:
                duplicate_groups.add(grupo)
            dedup[grupo] = row
        if duplicate_groups:
            warnings.append(f"Grupos repetidos no relatório de depreciação foram deduplicados sem somatório: {sorted(duplicate_groups)}.")
        rows = [dedup[g] for g in sorted(dedup)]

    if not rows:
        warnings.append("Nenhum grupo foi identificado no relatório de depreciação.")
    return ParsedDocument("depreciacao", filename, ug_code, ug_name, period_label, competence, len(pages), rows, warnings)


def identify_and_parse_pdf(file_obj: BinaryIO, filename: str, competence_hint: str | None = None) -> ParsedDocument:
    pages = _read_pdf_pages(file_obj)
    doc_type = detect_document_type(pages)
    file_obj.seek(0)
    if doc_type == "rsp":
        return parse_rsp(file_obj, filename)
    if doc_type == "depreciacao":
        return parse_depreciacao(file_obj, filename, competence_hint=competence_hint)
    raise ValueError(f"Não foi possível identificar o tipo do arquivo PDF: {filename}")
