# -*- coding: utf-8 -*-
"""Interface Streamlit para processamento de extrato FGTS TXT -> Excel."""

from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from fgts_extrato_to_excel import extrato_fgts_txt_para_excel


st.set_page_config(page_title="FGTS TXT para Excel", page_icon="📄", layout="wide")

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

:root {
  --bg: #020817;
  --card: #0f172a;
  --border: #1f2937;
  --muted: #94a3b8;
  --text: #e2e8f0;
  --primary: #3b82f6;
  --primary-hover: #2563eb;
}

html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
  background: var(--bg);
  color: var(--text);
  font-family: 'Inter', sans-serif;
}

.main .block-container {
  max-width: 1180px;
  padding-top: 2rem;
  padding-bottom: 2.5rem;
}

.dashboard-card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 18px;
  padding: 1.25rem;
  box-shadow: 0 10px 30px rgba(2, 8, 23, 0.35);
}

.header-title {
  font-size: 1.7rem;
  font-weight: 700;
  margin-bottom: 0.35rem;
}

.header-subtitle {
  color: var(--muted);
  font-size: 0.98rem;
  margin-bottom: 0;
}

[data-testid="stFileUploader"] {
  border: 1.6px dashed #334155;
  border-radius: 18px;
  background: rgba(30, 41, 59, 0.45);
  padding: 1.1rem;
}

[data-testid="stFileUploader"] section {
  min-height: 180px;
  display: flex;
  align-items: center;
  justify-content: center;
}

[data-testid="stFileUploader"] label {
  font-size: 1rem;
  color: var(--text);
}

.tip-card {
  background: rgba(59, 130, 246, 0.12);
  border: 1px solid rgba(59, 130, 246, 0.35);
  border-radius: 16px;
  padding: 1rem;
}

.tip-card h4 {
  margin-top: 0;
  margin-bottom: 0.55rem;
  color: #bfdbfe;
}

.tip-card p,
.tip-card li {
  color: #cbd5e1;
  font-size: 0.92rem;
}

[data-testid="stTextArea"] textarea {
  background: #0b1224;
  border: 1px solid var(--border);
  border-radius: 14px;
  color: var(--text);
  font-family: 'Fira Code', 'Consolas', monospace;
  min-height: 150px;
}

[data-testid="stTextInput"] input {
  background: #0b1224;
  border: 1px solid var(--border);
  border-radius: 12px;
  color: var(--text);
}

.stButton > button {
  border-radius: 12px;
  border: 1px solid var(--border);
  height: 44px;
  font-weight: 600;
}

.stButton > button[kind="primary"] {
  background: var(--primary);
  border-color: var(--primary);
  color: #eff6ff;
}

.stButton > button[kind="primary"]:hover {
  background: var(--primary-hover);
  border-color: var(--primary-hover);
}

.footer-actions {
  margin-top: 0.75rem;
}

.stDownloadButton button {
  border-radius: 12px;
  border: 1px solid #1d4ed8;
  background: rgba(37, 99, 235, 0.18);
  color: #dbeafe;
  font-weight: 600;
}
</style>
    """,
    unsafe_allow_html=True,
)

if "resultado_bytes" not in st.session_state:
    st.session_state.resultado_bytes = None
if "resultado_nome" not in st.session_state:
    st.session_state.resultado_nome = "Extrato_FGTS_Analitico_Processado.xlsx"
if "uploader_nonce" not in st.session_state:
    st.session_state.uploader_nonce = 0
if "output_name" not in st.session_state:
    st.session_state.output_name = "Extrato_FGTS_Analitico_Processado.xlsx"
if "dados_colados" not in st.session_state:
    st.session_state.dados_colados = ""

with st.container():
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.markdown('<p class="header-title">📊 Conversão de Extrato FGTS</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="header-subtitle">Transforme seu TXT em planilha XLSX com visual profissional e fluxo organizado.</p>',
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

col_upload, col_tip = st.columns([2, 1], gap="large")

with col_upload:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.subheader("Upload da planilha (principal)")
    st.caption("Arraste e solte o arquivo TXT ou clique para selecionar.")

    uploaded_file = st.file_uploader(
        "📎 Extrato analítico FGTS (.txt)",
        type=["txt"],
        key=f"uploaded_file_{st.session_state.uploader_nonce}",
        label_visibility="visible",
    )

    st.text_input(
        "Nome do arquivo de saída",
        key="output_name",
        help="Você pode alterar o nome do arquivo gerado.",
    )
    st.markdown("</div>", unsafe_allow_html=True)

with col_tip:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.markdown(
        """
<div class="tip-card">
  <h4>💡 Dica rápida</h4>
  <p>Para melhor resultado:</p>
  <ul>
    <li>Use o arquivo TXT original do extrato analítico.</li>
    <li>Mantenha o layout padrão da Caixa.</li>
    <li>Renomeie o arquivo final para facilitar auditorias.</li>
  </ul>
</div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
st.subheader("Dados colados (opcional)")
st.text_area(
    "Cole observações, blocos de referência ou dados auxiliares",
    key="dados_colados",
    placeholder="Ex.: observações internas, trechos para conferência, notas do processamento...",
    label_visibility="collapsed",
)
st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

st.markdown('<div class="dashboard-card footer-actions">', unsafe_allow_html=True)
act_col1, act_col2 = st.columns([1.2, 1], gap="small")

with act_col1:
    processar = st.button(
        "⚡ Processar extrato",
        type="primary",
        use_container_width=True,
        disabled=uploaded_file is None,
    )

with act_col2:
    limpar = st.button("Limpar campos", use_container_width=True)

if limpar:
    st.session_state.uploader_nonce += 1
    st.session_state.output_name = "Extrato_FGTS_Analitico_Processado.xlsx"
    st.session_state.dados_colados = ""
    st.session_state.resultado_bytes = None
    st.session_state.resultado_nome = "Extrato_FGTS_Analitico_Processado.xlsx"
    st.rerun()

if processar:
    if not uploaded_file:
        st.warning("Envie um arquivo TXT para continuar.")
    else:
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_path = Path(tmpdir)
                txt_path = tmp_path / uploaded_file.name
                txt_path.write_bytes(uploaded_file.getvalue())

                safe_output_name = st.session_state.output_name.strip() or "Extrato_FGTS_Analitico_Processado.xlsx"
                if not safe_output_name.lower().endswith(".xlsx"):
                    safe_output_name += ".xlsx"

                xlsx_path = tmp_path / safe_output_name
                extrato_fgts_txt_para_excel(str(txt_path), str(xlsx_path))

                st.session_state.resultado_bytes = xlsx_path.read_bytes()
                st.session_state.resultado_nome = safe_output_name

            st.success("Planilha gerada com sucesso!")

        except Exception as exc:  # interface deve retornar feedback amigável
            st.error(f"Não foi possível processar o arquivo: {exc}")

if st.session_state.resultado_bytes:
    st.download_button(
        "⬇️ Baixar planilha",
        data=st.session_state.resultado_bytes,
        file_name=st.session_state.resultado_nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("</div>", unsafe_allow_html=True)
