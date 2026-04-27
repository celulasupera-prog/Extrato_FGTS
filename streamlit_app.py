# -*- coding: utf-8 -*-
"""Interface Streamlit para processamento de extrato FGTS TXT -> Excel."""

from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from fgts_extrato_to_excel import extrato_fgts_txt_para_excel


st.set_page_config(
    page_title="FGTS TXT para Excel",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
  --bg: #020817;
  --surface: #0b1220;
  --surface-2: #0f172a;
  --surface-3: #111827;
  --border: rgba(148, 163, 184, .16);
  --border-strong: rgba(96, 165, 250, .28);
  --text: #f8fafc;
  --muted: #94a3b8;
  --muted-2: #cbd5e1;
  --primary: #3b82f6;
  --primary-2: #2563eb;
  --success: #22c55e;
  --danger: #ef4444;
  --radius-xl: 24px;
  --radius-lg: 18px;
  --shadow: 0 24px 70px rgba(0,0,0,.38);
}

html, body, [data-testid="stAppViewContainer"], .stApp {
  background:
    radial-gradient(circle at 12% 0%, rgba(59, 130, 246, .20), transparent 32rem),
    radial-gradient(circle at 90% 8%, rgba(14, 165, 233, .09), transparent 26rem),
    linear-gradient(180deg, #020817 0%, #030712 100%) !important;
  color: var(--text) !important;
  font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
}

[data-testid="stHeader"] { background: transparent !important; }
#MainMenu, footer { visibility: hidden; }

.main .block-container {
  max-width: 1180px;
  padding: 2.4rem 1.4rem 3rem;
}

h1, h2, h3, p, label, span, div { font-family: 'Inter', sans-serif !important; }
h1, h2, h3 { color: var(--text) !important; letter-spacing: -.025em; }
p, .stCaption, [data-testid="stCaptionContainer"] { color: var(--muted) !important; }

/* Cards por chave dos containers */
.st-key-upload_card,
.st-key-tip_card,
.st-key-actions_card,
.st-key-result_card {
  background: linear-gradient(180deg, rgba(15, 23, 42, .96), rgba(11, 18, 32, .96));
  border: 1px solid var(--border);
  border-radius: var(--radius-xl);
  padding: 1.25rem 1.25rem 1.35rem;
  box-shadow: var(--shadow);
}

.app-hero {
  position: relative;
  overflow: hidden;
  background:
    linear-gradient(135deg, rgba(15, 23, 42, .98), rgba(8, 13, 28, .96)),
    radial-gradient(circle at 90% 0%, rgba(59,130,246,.24), transparent 18rem);
  border: 1px solid var(--border);
  border-radius: 30px;
  padding: 30px;
  box-shadow: var(--shadow);
  margin-bottom: 22px;
}

.app-hero:after {
  content: "";
  position: absolute;
  inset: auto -12% -40% auto;
  width: 360px;
  height: 360px;
  border-radius: 999px;
  background: rgba(59, 130, 246, .09);
  filter: blur(8px);
}

.hero-grid {
  position: relative;
  z-index: 1;
  display: grid;
  grid-template-columns: 1fr 360px;
  gap: 24px;
  align-items: start;
}

.badge {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 8px 12px;
  border-radius: 999px;
  background: rgba(59,130,246,.12);
  border: 1px solid rgba(96,165,250,.25);
  color: #bfdbfe;
  font-size: 13px;
  font-weight: 700;
  margin-bottom: 15px;
}

.hero-title {
  font-size: clamp(2rem, 4vw, 3rem);
  line-height: 1.05;
  margin: 0 0 12px;
  color: #f8fafc !important;
  font-weight: 800;
}

.hero-subtitle {
  max-width: 740px;
  margin: 0;
  color: #cbd5e1 !important;
  font-size: 1rem;
  line-height: 1.65;
}

.metric-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 10px;
}

.metric-box {
  background: rgba(255,255,255,.035);
  border: 1px solid rgba(148,163,184,.13);
  border-radius: 18px;
  padding: 14px;
}

.metric-box small {
  display: block;
  color: var(--muted);
  font-size: 12px;
  margin-bottom: 5px;
}

.metric-box strong {
  color: #f8fafc;
  font-size: 18px;
}

.section-title {
  display: flex;
  align-items: center;
  gap: 10px;
  margin: 0 0 4px;
  color: var(--text) !important;
  font-size: 1.25rem;
  font-weight: 800;
}

.section-help {
  margin: 0 0 1rem;
  color: var(--muted) !important;
  font-size: .94rem;
  line-height: 1.55;
}

/* Upload premium */
[data-testid="stFileUploader"] {
  border: 1.5px dashed rgba(96,165,250,.42) !important;
  border-radius: 22px !important;
  background: linear-gradient(180deg, rgba(30,41,59,.48), rgba(15,23,42,.78)) !important;
  padding: 1.2rem !important;
  transition: all .2s ease;
}
[data-testid="stFileUploader"]:hover {
  border-color: rgba(96,165,250,.68) !important;
  background: linear-gradient(180deg, rgba(30,41,59,.64), rgba(15,23,42,.92)) !important;
}
[data-testid="stFileUploader"] section {
  min-height: 185px;
  display: flex;
  align-items: center;
  justify-content: center;
}
[data-testid="stFileUploader"] small { color: var(--muted) !important; }
[data-testid="stFileUploader"] button {
  border-radius: 12px !important;
  border: 1px solid rgba(96,165,250,.35) !important;
  background: linear-gradient(135deg, rgba(59,130,246,.22), rgba(37,99,235,.28)) !important;
  color: #eff6ff !important;
  font-weight: 800 !important;
  box-shadow: 0 8px 18px rgba(37,99,235,.22) !important;
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  gap: .3rem !important;
  white-space: nowrap !important;
  overflow: hidden !important;
}
[data-testid="stFileUploader"] button:hover {
  border-color: rgba(147,197,253,.8) !important;
  background: linear-gradient(135deg, rgba(59,130,246,.34), rgba(37,99,235,.46)) !important;
  color: #ffffff !important;
}
[data-testid="stFileUploader"] button p {
  margin: 0 !important;
  line-height: 1.1 !important;
}

/* Inputs */
[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea {
  background: rgba(2, 6, 23, .56) !important;
  border: 1px solid rgba(148,163,184,.16) !important;
  border-radius: 16px !important;
  color: var(--text) !important;
  box-shadow: inset 0 1px 0 rgba(255,255,255,.03) !important;
}
[data-testid="stTextInput"] input:focus,
[data-testid="stTextArea"] textarea:focus {
  border-color: rgba(96,165,250,.62) !important;
  box-shadow: 0 0 0 3px rgba(59,130,246,.14) !important;
}
[data-testid="stTextArea"] textarea {
  min-height: 220px !important;
  font-family: 'Consolas', 'Fira Code', ui-monospace, monospace !important;
  font-size: .94rem !important;
  line-height: 1.7 !important;
}

.tip-panel {
  border-radius: 20px;
  padding: 18px;
  background: rgba(59, 130, 246, .105);
  border: 1px solid rgba(96, 165, 250, .24);
}
.tip-panel h3 {
  margin: 0 0 10px;
  font-size: 1.05rem;
  color: #dbeafe !important;
}
.tip-panel p, .tip-panel li {
  color: #cbd5e1 !important;
  font-size: .92rem;
  line-height: 1.65;
}
.tip-panel ul { margin: 12px 0 0; padding-left: 20px; }

.status-row {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 10px;
  margin-top: 14px;
}
.status-pill {
  border-radius: 16px;
  padding: 12px;
  background: rgba(2,6,23,.32);
  border: 1px solid rgba(148,163,184,.12);
}
.status-pill small { color: var(--muted); display: block; font-size: 12px; margin-bottom: 4px; }
.status-pill b { color: var(--text); font-size: 14px; }

/* Botões */
.stButton > button,
.stDownloadButton > button {
  height: 48px !important;
  border-radius: 15px !important;
  font-weight: 800 !important;
  border: 1px solid rgba(148,163,184,.18) !important;
  transition: all .18s ease !important;
}
.stButton > button:hover,
.stDownloadButton > button:hover {
  transform: translateY(-1px);
}
.stButton > button[kind="primary"] {
  background: linear-gradient(135deg, var(--primary), var(--primary-2)) !important;
  border-color: rgba(96,165,250,.55) !important;
  color: white !important;
  box-shadow: 0 16px 34px rgba(37,99,235,.32) !important;
}
.stDownloadButton > button {
  background: rgba(34,197,94,.14) !important;
  border-color: rgba(34,197,94,.34) !important;
  color: #bbf7d0 !important;
}

/* Alertas nativos */
[data-testid="stAlert"] {
  border-radius: 16px !important;
  border: 1px solid rgba(148,163,184,.16) !important;
}

.footer-note {
  color: var(--muted) !important;
  font-size: .9rem;
  margin: .5rem 0 0;
}

@media (max-width: 900px) {
  .hero-grid { grid-template-columns: 1fr; }
  .metric-grid, .status-row { grid-template-columns: 1fr; }
  .app-hero { padding: 22px; border-radius: 24px; }
  .main .block-container { padding-left: 1rem; padding-right: 1rem; }
}
</style>
""",
    unsafe_allow_html=True,
)

DEFAULT_OUTPUT_NAME = "Extrato_FGTS_Analitico_Processado.xlsx"

if "resultado_bytes" not in st.session_state:
    st.session_state.resultado_bytes = None
if "resultado_nome" not in st.session_state:
    st.session_state.resultado_nome = DEFAULT_OUTPUT_NAME
if "uploader_nonce" not in st.session_state:
    st.session_state.uploader_nonce = 0
if "output_name" not in st.session_state:
    st.session_state.output_name = DEFAULT_OUTPUT_NAME
st.markdown(
    """
<section class="app-hero">
  <div class="hero-grid">
    <div>
      <div class="badge">📄 FGTS • Conversão inteligente</div>
      <h1 class="hero-title">Extrato FGTS TXT para Excel</h1>
      <p class="hero-subtitle">
        Envie o extrato analítico em TXT, processe os dados automaticamente e baixe uma planilha XLSX pronta para conferência, auditoria e envio.
      </p>
    </div>
    <div class="metric-grid">
      <div class="metric-box"><small>Entrada</small><strong>TXT</strong></div>
      <div class="metric-box"><small>Saída</small><strong>XLSX</strong></div>
      <div class="metric-box"><small>Fluxo</small><strong>Rápido</strong></div>
    </div>
  </div>
</section>
""",
    unsafe_allow_html=True,
)

col_upload, col_tip = st.columns([1.55, 1], gap="large")

with col_upload:
    with st.container(key="upload_card"):
        st.markdown('<h2 class="section-title">⬆️ Upload do extrato</h2>', unsafe_allow_html=True)
        st.markdown(
            '<p class="section-help">Selecione o arquivo TXT original do extrato analítico FGTS. Depois ajuste o nome do arquivo final, se necessário.</p>',
            unsafe_allow_html=True,
        )

        uploaded_file = st.file_uploader(
            "Arraste o arquivo TXT aqui ou clique para selecionar",
            type=["txt"],
            key=f"uploaded_file_{st.session_state.uploader_nonce}",
            label_visibility="visible",
        )

        st.text_input(
            "Nome do arquivo de saída",
            key="output_name",
            help="O arquivo será salvo em .xlsx. Se você não informar a extensão, ela será adicionada automaticamente.",
        )

with col_tip:
    with st.container(key="tip_card"):
        st.markdown(
            """
<div class="tip-panel">
  <h3>💡 Antes de processar</h3>
  <p>Para evitar erro na conversão, use o arquivo TXT original e mantenha o layout padrão gerado pela Caixa.</p>
  <ul>
    <li>Não edite o TXT manualmente.</li>
    <li>Evite arquivos copiados de PDF.</li>
    <li>Renomeie a saída para facilitar auditoria.</li>
  </ul>
</div>
<div class="status-row">
  <div class="status-pill"><small>Status</small><b>Pronto</b></div>
  <div class="status-pill"><small>Arquivo</small><b>Obrigatório</b></div>
  <div class="status-pill"><small>Saída</small><b>XLSX</b></div>
</div>
""",
            unsafe_allow_html=True,
        )

st.markdown("<div style='height: 18px'></div>", unsafe_allow_html=True)

with st.container(key="actions_card"):
    action_col_1, action_col_2 = st.columns([1.4, 1], gap="small")

    with action_col_1:
        processar = st.button(
            "⚡ Processar extrato",
            type="primary",
            use_container_width=True,
            disabled=uploaded_file is None,
        )

    with action_col_2:
        limpar = st.button("🧹 Limpar campos", use_container_width=True)

    st.markdown(
        '<p class="footer-note">Após o processamento, o botão de download aparecerá logo abaixo.</p>',
        unsafe_allow_html=True,
    )

if limpar:
    st.session_state.uploader_nonce += 1
    st.session_state.output_name = DEFAULT_OUTPUT_NAME
    st.session_state.resultado_bytes = None
    st.session_state.resultado_nome = DEFAULT_OUTPUT_NAME
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

                safe_output_name = st.session_state.output_name.strip() or DEFAULT_OUTPUT_NAME
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
    st.markdown("<div style='height: 18px'></div>", unsafe_allow_html=True)
    with st.container(key="result_card"):
        st.markdown('<h2 class="section-title">✅ Arquivo pronto</h2>', unsafe_allow_html=True)
        st.markdown(
            f'<p class="section-help">O arquivo <strong>{st.session_state.resultado_nome}</strong> foi gerado com sucesso.</p>',
            unsafe_allow_html=True,
        )
        st.download_button(
            "⬇️ Baixar planilha XLSX",
            data=st.session_state.resultado_bytes,
            file_name=st.session_state.resultado_nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
