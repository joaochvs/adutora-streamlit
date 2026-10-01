import streamlit as st
from core import ler_excel, preparar, gerar_pdf

st.set_page_config(page_title="Revisitas de campo", page_icon="📋")
st.title("Revisitas de campo")
st.write("Envie o Excel e baixe o PDF dos ausentes com o próximo turno sugerido.")
upload = st.file_uploader("Planilha de visitas", type=["xlsx", "xlsm"])
if upload is not None:
    try:
        with st.spinner("Gerando relatório..."):
            df, sheet = ler_excel(upload.getvalue())
            result = preparar(df)
            pdf = gerar_pdf(result, f"{upload.name} / {sheet}")
        st.success(f"Relatório pronto: {len(result)} registro(s) com status Ausente.")
        st.download_button("Baixar PDF", pdf, "relatorio_revisitas.pdf", "application/pdf", type="primary")
    except Exception as exc:
        st.error(f"Não foi possível gerar o PDF: {exc}")
