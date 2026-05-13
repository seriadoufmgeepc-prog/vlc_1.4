from __future__ import annotations

import base64
import io
from pathlib import Path
from typing import Iterable

import pandas as pd
import streamlit as st

from src.parsers import identify_and_parse_pdf
from src.quality import analyse_results, totals_dict
from src.reporting import build_outputs, build_siafi_web_df, generate_excel, generate_pdf, load_pcasp_map
from src.utils import file_signature, format_br_number, format_brasilia_datetime, safe_filename_part

BASE_DIR = Path(__file__).resolve().parent
PCASP_PATH = BASE_DIR / "modelo_grupo_x_pcasp.xlsx"
LOGO_PATH = BASE_DIR / "proplan_ufmg.jpg"


# -----------------------------------------------------------------------------
# Camada de interface
# -----------------------------------------------------------------------------

def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root { --azul:#1F4E78; --azul-escuro:#173B5C; --vermelho:#C8102E; --borda:#DDE4EC; --cinza:#667085; --fundo-card:#FFFFFF; --fundo-painel:#F8FAFC; }
        .stApp { background: linear-gradient(180deg,#F7F9FC 0%,#F2F5F9 100%); }
        .block-container { max-width: 1320px; padding-top: 1.9rem; padding-bottom: 1.5rem; padding-left: 2rem; padding-right: 2rem; }
        .app-header-shell { margin: .25rem 0 1rem 0; }
        .app-header { display:flex; justify-content:space-between; gap:1.40rem; align-items:center; padding:1.10rem 1.30rem; background:#fff; border:1px solid var(--borda); border-left:6px solid var(--vermelho); border-radius:18px; box-shadow:0 10px 24px rgba(31,78,120,.075); }
        .app-text { min-width: 0; flex:1 1 auto; }
        .eyebrow { color:var(--azul); font-size:.82rem; font-weight:760; letter-spacing:.015em; margin-bottom:.20rem; line-height:1.15; text-transform:uppercase; }
        .app-text h1 { margin:.03rem 0 0 0; color:#101828; font-size:clamp(1.74rem,2.35vw,2.18rem); font-weight:780; line-height:1.06; }
        .subtitle { margin-top:.18rem; color:#23313f; font-size:1.04rem; font-weight:660; }
        .flowline { margin-top:.22rem; color:#5E6B78; font-size:.88rem; line-height:1.38; max-width: 95%; }
        .app-side { display:flex; flex-direction:column; align-items:center; justify-content:center; gap:.55rem; min-width: 250px; }
        .logo-frame { display:flex; align-items:center; justify-content:center; width:100%; padding:.45rem .90rem; border-radius:16px; background:linear-gradient(180deg,#FFFFFF 0%, #F8FAFC 100%); border:1px solid #E6ECF3; }
        .app-logo { width:138px; max-height:84px; object-fit:contain; }
        .dt-badge { color:#475467; background:#F3F6F9; border:1px solid #E2E8F0; border-radius:999px; padding:.28rem .68rem; font-size:.73rem; font-weight:650; white-space:nowrap; }
        .panel-card { background:#fff; border:1px solid var(--borda); border-radius:14px; padding:1rem 1.1rem; box-shadow:0 5px 14px rgba(15,23,42,.045); margin:.55rem 0 .85rem 0; }
        .panel-title { font-weight:750; color:#101828; font-size:1.02rem; margin-bottom:.30rem; }
        .panel-caption { color:#667085; font-size:.88rem; line-height:1.42; }
        div[data-testid="stFileUploader"], div[data-testid="stExpander"] { background:#fff; border:1px solid var(--borda); border-radius:12px; box-shadow:0 4px 14px rgba(15,23,42,.04); }
        div[data-testid="stFileUploader"] { padding:.72rem .85rem; }
        .stButton button, .stDownloadButton button { border-radius:10px; padding:.52rem .96rem; font-weight:680; }
        .stButton button[kind="primary"], .stDownloadButton button { background:var(--azul); color:#fff; border:1px solid var(--azul); }
        .stButton button:hover, .stDownloadButton button:hover { background:var(--azul-escuro); color:#fff; border-color:var(--azul-escuro); }
        .stAlert { border-radius:10px; }
        div[data-testid="stMetric"] { background:#fff; border:1px solid var(--borda); border-radius:12px; padding:.75rem .8rem; box-shadow:0 4px 12px rgba(15,23,42,.035); }
        .result-marker { padding:.65rem .85rem; border-radius:12px; background:#F0F7FF; border:1px solid #C7DFF6; color:#173B5C; font-weight:650; margin:.65rem 0; }
        .nav-caption { margin:.55rem 0 .20rem 0; color:#475467; font-size:.82rem; font-weight:650; }
        div[data-testid="stRadio"] > div { gap:.40rem; }
        div[data-testid="stRadio"] [role="radiogroup"] { display:flex; flex-wrap:wrap; gap:.55rem; padding:.10rem 0 .55rem 0; }
        div[data-testid="stRadio"] label { margin:0 !important; }
        div[data-testid="stRadio"] [data-testid="stMarkdownContainer"] p { margin:0; }
        div[data-testid="stRadio"] [role="radio"] { background:#fff; border:1px solid var(--borda); border-radius:12px; padding:.60rem .92rem; box-shadow:0 4px 12px rgba(15,23,42,.04); min-height:44px; display:flex; align-items:center; }
        div[data-testid="stRadio"] [role="radio"][aria-checked="true"] { background:linear-gradient(180deg,#EDF5FF 0%, #E2EEFB 100%); border-color:#B8D2F0; box-shadow:0 6px 16px rgba(31,78,120,.10); }
        div[data-testid="stRadio"] [role="radio"]:focus-visible { outline:2px solid #7CB5EC; outline-offset:2px; }
        .section-head { margin:.10rem 0 .65rem 0; color:#101828; font-size:1.06rem; font-weight:760; }
        @media (max-width: 900px){ .app-header{flex-direction:column;align-items:flex-start;} .app-side{align-items:flex-start; min-width: unset; width:100%;} .logo-frame{justify-content:flex-start;} .flowline{max-width:100%;} }
        @media (max-width: 780px){ .block-container { padding-left:1rem; padding-right:1rem; padding-top:1.4rem; } .app-logo{width:126px; max-height:78px;} }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    logo_html = ""
    if LOGO_PATH.exists():
        logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode("utf-8")
        logo_html = f'<div class="logo-frame"><img class="app-logo" src="data:image/jpeg;base64,{logo_b64}" alt="PROPLAN/UFMG" /></div>'
    st.markdown(
        f"""
        <div class="app-header-shell">
          <div class="app-header">
            <div class="app-text">
              <div class="eyebrow">Universidade Federal de Minas Gerais • PROPLAN • Divisão de Contabilidade</div>
              <h1>SADPat — Sistema Auxiliar de Desfazimento Patrimonial</h1>
              <div class="subtitle">Apuração contábil auxiliar do desfazimento patrimonial</div>
              <div class="flowline">Relatório Sintético Patrimonial de Saídas, com exclusão das saídas por transferência, + Relatório de Depreciação Acumulada → Relatório Sintético, Registro SIAFI Web, Excel e PDF.</div>
            </div>
            <div class="app-side">{logo_html}<span class="dt-badge">Horário de Brasília: {format_brasilia_datetime()}</span></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_intro_card() -> None:
    st.markdown(
        """
        <div class="panel-card">
          <div class="panel-title">📂 Entrada única dos relatórios</div>
          <div class="panel-caption">Anexe os dois PDFs da mesma competência. O aplicativo identifica automaticamente o Relatório Sintético Patrimonial e o Relatório de Depreciação Acumulada, desconsidera as saídas por transferência, cruza os grupos, aplica o mapeamento PCASP interno e mantém os resultados visíveis após o processamento.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# Estado, classificação e processamento
# -----------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def parse_pdf_payload(file_bytes: bytes, filename: str, competence_hint: str | None = None, parser_version: str = "v13_doacao_imb037"):
    """Parseia o PDF com cache e invalida resultados de versões anteriores do parser."""
    # parser_version é argumento deliberado para evitar reutilização de cache
    # de versões anteriores quando o app é atualizado no Streamlit.
    _ = parser_version
    return identify_and_parse_pdf(io.BytesIO(file_bytes), filename, competence_hint=competence_hint)


def uploaded_signature(uploaded_files: Iterable[object]) -> str:
    parts: list[str] = []
    for file in uploaded_files:
        size = getattr(file, "size", None)
        parts.append(f"{getattr(file, 'name', '')}:{size}")
    return file_signature(parts)


def clear_results_if_files_changed(uploaded_files: list[object]) -> None:
    sig = uploaded_signature(uploaded_files) if uploaded_files else ""
    prev = st.session_state.get("uploaded_sig", "")
    if sig != prev:
        st.session_state.pop("result_payload", None)
        st.session_state["uploaded_sig"] = sig


def classify_uploaded_files(uploaded_files) -> tuple[object | None, object | None, list[str]]:
    rsp_file = None
    dep_file = None
    errors: list[str] = []
    for uploaded in uploaded_files:
        try:
            parsed = parse_pdf_payload(uploaded.getvalue(), uploaded.name)
            uploaded.seek(0)
            if parsed.doc_type == "rsp":
                if rsp_file is not None:
                    errors.append("Foi identificado mais de um Relatório Sintético Patrimonial. Mantenha apenas um PDF desse tipo.")
                rsp_file = uploaded
            elif parsed.doc_type == "depreciacao":
                if dep_file is not None:
                    errors.append("Foi identificado mais de um Relatório de Depreciação Acumulada. Mantenha apenas um PDF desse tipo.")
                dep_file = uploaded
        except Exception as exc:
            errors.append(f"Falha ao identificar o arquivo {uploaded.name}: {exc}")
    return rsp_file, dep_file, errors


def process_files(rsp_file, dep_file) -> dict:
    rsp_parsed = parse_pdf_payload(rsp_file.getvalue(), rsp_file.name)
    dep_parsed = parse_pdf_payload(dep_file.getvalue(), dep_file.name, competence_hint=rsp_parsed.competence)
    rsp_df = rsp_parsed.to_frame()
    dep_df = dep_parsed.to_frame()

    if rsp_df.empty:
        raise ValueError("Nenhuma linha válida foi extraída do Relatório Sintético Patrimonial.")
    if dep_df.empty:
        raise ValueError("Nenhuma linha válida foi extraída do Relatório de Depreciação Acumulada.")
    if not PCASP_PATH.exists():
        raise FileNotFoundError("Base interna modelo_grupo_x_pcasp.xlsx não localizada no projeto.")

    pcasp_df = load_pcasp_map(PCASP_PATH)
    final_df, memoria_df, logs_df, meta_df = build_outputs(
        rsp_df,
        dep_df,
        pcasp_df,
        rsp_parsed.to_meta() | {"filename": rsp_parsed.filename},
        dep_parsed.to_meta() | {"filename": dep_parsed.filename},
    )
    header = {
        "ug_code": rsp_parsed.ug_code or dep_parsed.ug_code or "",
        "ug_name": rsp_parsed.ug_name or dep_parsed.ug_name or "",
        "competence": rsp_parsed.competence or dep_parsed.competence or "",
    }
    analysis_df = analyse_results(final_df, memoria_df)
    siafi_df = build_siafi_web_df(final_df)
    excel_bytes = generate_excel(final_df, memoria_df, logs_df, meta_df, header)
    pdf_bytes = generate_pdf(final_df, header, logs_df, meta_df)
    return {
        "final_df": final_df,
        "memoria_df": memoria_df,
        "siafi_df": siafi_df,
        "logs_df": logs_df,
        "meta_df": meta_df,
        "analysis_df": analysis_df,
        "excel_bytes": excel_bytes,
        "pdf_bytes": pdf_bytes,
        "warnings": rsp_parsed.warnings + dep_parsed.warnings,
        "header": header,
        "processed_at": format_brasilia_datetime(),
    }


# -----------------------------------------------------------------------------
# Exibição dos resultados
# -----------------------------------------------------------------------------

def format_display_df(df: pd.DataFrame, include_total: bool = False) -> pd.DataFrame:
    out = df.copy()
    if include_total and not out.empty:
        total_row = {col: "" for col in out.columns}
        total_row["Grupo"] = "TOTAL"
        for col in ["Saídas/Grupo", "Dep. Acumulada", "Vlr. Liq. Contábil"]:
            if col in out.columns:
                total_row[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).sum()
        out = pd.concat([out, pd.DataFrame([total_row])], ignore_index=True)

    if "Grupo" in out.columns:
        def fmt_group(x):
            if pd.isna(x) or x == "":
                return ""
            if str(x).upper() == "TOTAL":
                return "TOTAL"
            try:
                return str(int(float(x)))
            except Exception:
                return str(x)
        out["Grupo"] = out["Grupo"].map(fmt_group)

    for col in out.columns:
        if col == "Grupo":
            continue
        if pd.api.types.is_numeric_dtype(out[col]) or col in ["Saídas/Grupo", "Dep. Acumulada", "Vlr. Liq. Contábil", "Situação IMB010", "Situação IMB025", "Situação IMB037"]:
            out[col] = out[col].map(lambda x: format_br_number(x, 2) if pd.notna(x) and x != "" else "")
    return out




def right_align_second_third_columns(df: pd.DataFrame):
    """Aplica alinhamento visual à direita nas colunas 2 e 3 do corpo da tabela.

    O ajuste é usado apenas na exibição do quadro Logs e Metadados, preservando
    cabeçalhos centralizados e sem alterar os dados originais.
    """
    if df is None or df.empty:
        return df
    target_cols = [col for idx, col in enumerate(df.columns, start=1) if idx in {2, 3}]
    if not target_cols:
        return df
    return (
        df.style
        .set_properties(subset=target_cols, **{"text-align": "right"})
        .set_table_styles([{"selector": "th", "props": [("text-align", "center")]}])
    )

def render_downloads(payload: dict) -> None:
    header = payload["header"]
    base_name = f"relatorio_de_baixas_{safe_filename_part(header.get('ug_code'))}_{safe_filename_part(header.get('competence'))}"
    c1, c2, c3 = st.columns([1.15, 1.65, 2.2])
    with c1:
        st.download_button(
            "⬇️ Baixar Excel",
            data=payload["excel_bytes"],
            file_name=f"{base_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with c2:
        st.download_button(
            "⬇️ Exportar Relatório de Baixas em PDF",
            data=payload["pdf_bytes"],
            file_name=f"{base_name}.pdf",
            mime="application/pdf",
        )
    with c3:
        st.markdown(f"<div class='result-marker'>Resultados preservados na tela • Processado em {payload.get('processed_at', '')}</div>", unsafe_allow_html=True)


def render_metrics(final_df: pd.DataFrame) -> None:
    totals = totals_dict(final_df)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Grupos apurados", len(final_df))
    c2.metric("Demais Saídas", f"R$ {format_br_number(totals['Saídas/Grupo'])}")
    c3.metric("Dep. Acumulada", f"R$ {format_br_number(totals['Dep. Acumulada'])}")
    c4.metric("Vlr. Liq. Contábil", f"R$ {format_br_number(totals['Vlr. Liq. Contábil'])}")


def render_results(payload: dict) -> None:
    final_df = payload["final_df"]
    memoria_df = payload["memoria_df"]
    siafi_df = payload.get("siafi_df", build_siafi_web_df(final_df))
    logs_df = payload["logs_df"]
    meta_df = payload["meta_df"]
    analysis_df = payload["analysis_df"]

    render_downloads(payload)
    render_metrics(final_df)

    st.markdown("<div class='nav-caption'>Navegação dos resultados</div>", unsafe_allow_html=True)
    sections = {
        "📊 Relatório Sintético": "final",
        "🧾 Registro SIAFI Web": "siafi",
        "✅ Análise de Consistência": "analysis",
        "🧠 Memória de Cálculo": "memory",
        "📝 Logs e Metadados": "logs",
    }
    selected_label = st.radio(
        "Navegação dos resultados",
        options=list(sections.keys()),
        index=list(sections.values()).index(st.session_state.get("results_nav", "final")) if st.session_state.get("results_nav", "final") in sections.values() else 0,
        horizontal=True,
        label_visibility="collapsed",
        key="results_nav_radio",
    )
    current_section = sections[selected_label]
    st.session_state["results_nav"] = current_section

    if current_section == "final":
        st.markdown("<div class='section-head'>📊 Relatório Sintético</div>", unsafe_allow_html=True)
        st.dataframe(format_display_df(final_df, include_total=True), use_container_width=True, hide_index=True)

    elif current_section == "siafi":
        st.markdown("<div class='section-head'>🧾 Registro SIAFI Web</div>", unsafe_allow_html=True)
        st.dataframe(format_display_df(siafi_df), use_container_width=True, hide_index=True)

    elif current_section == "analysis":
        st.markdown("<div class='section-head'>✅ Análise dos resultados apresentados</div>", unsafe_allow_html=True)
        st.dataframe(analysis_df, use_container_width=True, hide_index=True)
        if payload["warnings"]:
            with st.expander("⚠️ Avisos de leitura dos PDFs", expanded=False):
                for item in payload["warnings"]:
                    st.write(f"- {item}")

    elif current_section == "memory":
        st.markdown("<div class='section-head'>🧠 Memória de Cálculo</div>", unsafe_allow_html=True)
        st.dataframe(format_display_df(memoria_df), use_container_width=True, hide_index=True)

    elif current_section == "logs":
        left, right = st.columns(2)
        with left:
            st.markdown("<div class='section-head'>📝 Logs</div>", unsafe_allow_html=True)
            st.dataframe(right_align_second_third_columns(logs_df), use_container_width=True, hide_index=True)
        with right:
            st.markdown("<div class='section-head'>🗂️ Metadados</div>", unsafe_allow_html=True)
            st.dataframe(right_align_second_third_columns(meta_df), use_container_width=True, hide_index=True)


# -----------------------------------------------------------------------------
# Entrada principal
# -----------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(page_title="SADPat — Sistema Auxiliar de Desfazimento Patrimonial", layout="wide")
    inject_css()
    render_header()
    render_intro_card()

    with st.expander("ℹ️ Orientações de uso", expanded=False):
        st.write(
            "1. Anexe o Relatório Sintético Patrimonial de Saídas e o Relatório de Depreciação Acumulada. O sistema considera somente as demais saídas, excluindo saídas por transferência. "
            "2. Clique em Gerar relatório de apuração. "
            "3. Confira o Relatório Sintético, o Registro SIAFI Web, os Logs e Metadados antes de baixar os arquivos. "
            "4. Para novo processamento, remova ou substitua os PDFs carregados."
        )

    uploaded_files = st.file_uploader(
        "📤 Anexar relatórios em PDF",
        type=["pdf"],
        accept_multiple_files=True,
        help="Envie exatamente dois PDFs: Relatório Sintético Patrimonial e Relatório de Depreciação Acumulada.",
    )
    uploaded_files = uploaded_files or []
    clear_results_if_files_changed(uploaded_files)

    if not uploaded_files:
        st.info("Aguardando o envio dos arquivos para processamento.")
        return

    rsp_file, dep_file, errors = classify_uploaded_files(uploaded_files)
    for msg in errors:
        st.warning(msg)
    if rsp_file is None or dep_file is None:
        st.error("Envie exatamente um Relatório Sintético Patrimonial em PDF e um Relatório de Depreciação Acumulada em PDF.")
        return

    action_col1, action_col2 = st.columns([1.35, 3.65])
    with action_col1:
        if st.button("▶️ Gerar relatório de apuração", type="primary"):
            try:
                st.session_state["result_payload"] = process_files(rsp_file, dep_file)
                st.success("Processamento concluído com sucesso.")
            except Exception as exc:
                st.session_state.pop("result_payload", None)
                st.error(f"Falha no processamento: {exc}")
    with action_col2:
        st.caption("Os resultados permanecem visíveis após a geração. Para um novo ciclo, remova ou substitua os arquivos carregados.")

    payload = st.session_state.get("result_payload")
    if payload:
        render_results(payload)


if __name__ == "__main__":
    main()
