import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Remover Duplicados Excel", layout="centered")

st.title("🧹 Remover Duplicados de Excel")
st.write("Faça upload do arquivo, escolha a coluna base e gere um novo Excel sem duplicados.")

# ===== UPLOAD DO ARQUIVO =====
arquivo = st.file_uploader(
    "Selecione o arquivo Excel",
    type=["xlsx", "xls"]
)

if arquivo:
    df = pd.read_excel(arquivo)
    st.success("Arquivo carregado com sucesso!")

    st.write("### Pré-visualização")
    st.dataframe(df.head())

    # ===== ESCOLHER COLUNA =====
    coluna_base = st.selectbox(
        "Selecione a coluna para remover duplicados:",
        df.columns
    )

    # ===== OPÇÕES EXTRAS =====
    st.write("### Opções")
    normalizar = st.checkbox("Ignorar maiúsculas, espaços e variações de texto", value=True)

    if st.button("🚀 Gerar arquivo sem duplicados"):
        df_trabalho = df.copy()

        if normalizar:
            df_trabalho[coluna_base] = (
                df_trabalho[coluna_base]
                .astype(str)
                .str.lower()
                .str.strip()
                .str.replace(r"\s+", " ", regex=True)
            )

        total_linhas = len(df_trabalho)
        df_resultado = df_trabalho.drop_duplicates(subset=[coluna_base], keep="first")

        # ===== GERAR ARQUIVO =====
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_saida = f"arquivo_sem_duplicados_{timestamp}.xlsx"

        df_resultado.to_excel(nome_saida, index=False)

        st.success("Arquivo gerado com sucesso!")

        st.write("### 📊 Resumo")
        st.write(f"• Coluna usada: **{coluna_base}**")
        st.write(f"• Linhas originais: **{total_linhas}**")
        st.write(f"• Linhas finais: **{len(df_resultado)}**")

        with open(nome_saida, "rb") as file:
            st.download_button(
                label="⬇️ Baixar arquivo Excel",
                data=file,
                file_name=nome_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
