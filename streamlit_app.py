# -*- coding: utf-8 -*-
"""Interface Streamlit para processamento de extrato FGTS TXT -> Excel."""

from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from fgts_extrato_to_excel import extrato_fgts_txt_para_excel


st.set_page_config(page_title="FGTS TXT para Excel", page_icon="📄", layout="centered")

st.title("📄 FGTS TXT para Excel")
st.write(
    "Faça upload do extrato analítico FGTS em TXT e baixe a planilha processada em XLSX "
    "(uma aba por trabalhador)."
)

uploaded_file = st.file_uploader("Selecione o arquivo TXT do extrato", type=["txt"])
output_name = st.text_input(
    "Nome do arquivo de saída",
    value="Extrato_FGTS_Analitico_Processado.xlsx",
    help="Você pode alterar o nome do arquivo gerado.",
)

if st.button("Processar extrato", type="primary", disabled=uploaded_file is None):
    if not uploaded_file:
        st.warning("Envie um arquivo TXT para continuar.")
    else:
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_path = Path(tmpdir)
                txt_path = tmp_path / uploaded_file.name
                txt_path.write_bytes(uploaded_file.getvalue())

                safe_output_name = output_name.strip() or "Extrato_FGTS_Analitico_Processado.xlsx"
                if not safe_output_name.lower().endswith(".xlsx"):
                    safe_output_name += ".xlsx"

                xlsx_path = tmp_path / safe_output_name
                extrato_fgts_txt_para_excel(str(txt_path), str(xlsx_path))

                result_bytes = xlsx_path.read_bytes()

            st.success("Planilha gerada com sucesso!")
            st.download_button(
                "⬇️ Baixar planilha",
                data=result_bytes,
                file_name=safe_output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as exc:  # interface deve retornar feedback amigável
            st.error(f"Não foi possível processar o arquivo: {exc}")

with st.expander("Como usar"):
    st.markdown(
        """
1. Clique em **Browse files** e envie seu arquivo `.txt` do extrato analítico FGTS.
2. (Opcional) Ajuste o nome do arquivo de saída.
3. Clique em **Processar extrato**.
4. Faça download do `.xlsx` gerado.
        """
    )
