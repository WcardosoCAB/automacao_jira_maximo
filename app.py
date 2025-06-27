
import streamlit as st
import pandas as pd
import tempfile
import os
from io import BytesIO

from seu_codigo import ler_jira, ler_maximo, aplicar_formatacao_excel, verificar_sistemas_em_fechamento

st.set_page_config(page_title="Gerador de Planilhas Jira/Maximo", layout="centered")
st.title("📊 Gerador de Planilhas Jira/Maximo")
st.write("Envie as planilhas extraídas do Jira e do Maximo para gerar o Excel padronizado.")

jira_file = st.file_uploader("📁 Envie o arquivo Jira.xlsx", type=["xlsx"])
maximo_file = st.file_uploader("📁 Envie o arquivo Maximo.xlsx ou Maximo.csv", type=["xlsx", "csv"])

if jira_file and maximo_file:
    with tempfile.TemporaryDirectory() as tmpdirname:
        caminho_jira = os.path.join(tmpdirname, "Jira.xlsx")
        with open(caminho_jira, "wb") as f:
            f.write(jira_file.read())

        caminho_maximo = os.path.join(tmpdirname, maximo_file.name)
        with open(caminho_maximo, "wb") as f:
            f.write(maximo_file.read())

        colunas_finais = [
            "Chave", "Resumo", "Status", "Descrição", "Relator",
            "Planned start date", "Planned end date"
        ]

        try:
            df_jira = ler_jira(caminho_jira, colunas_finais)
            df_maximo = ler_maximo(tmpdirname, colunas_finais)

            caminho_saida = os.path.join(tmpdirname, "planilha_final.xlsx")
            abas_criadas = []

            with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
                if df_jira is not None:
                    df_jira.to_excel(writer, sheet_name='Jira', index=False)
                    abas_criadas.append('Jira')

                df_maximo.to_excel(writer, sheet_name='Maximo', index=False)
                abas_criadas.append('Maximo')

                df_dummy = pd.DataFrame([[""]])
                df_dummy.to_excel(writer, sheet_name='Participantes', index=False, header=False)
                abas_criadas.append('Participantes')

            aplicar_formatacao_excel(caminho_saida, abas_criadas)
            verificar_sistemas_em_fechamento(df_jira, df_maximo, caminho_saida)

            with open(caminho_saida, "rb") as f:
                bytes_data = f.read()
                st.success("✅ Planilha gerada com sucesso!")
                st.download_button(
                    label="📥 Baixar Planilha Final",
                    data=bytes_data,
                    file_name="planilha_final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Erro ao processar os arquivos: {e}")
else:
    st.info("Por favor, envie os dois arquivos para iniciar o processamento.")
