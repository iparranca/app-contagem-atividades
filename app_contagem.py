import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def definir_ano_letivo(data):
    ano = data.year
    if data.month >= 9:
        return f"{ano}/{ano + 1}"
    elif data.month <= 7:
        return f"{ano - 1}/{ano}"
    else:
        return "Desconhecido"

st.title("Contagem de Atividades por Ano Letivo")

uploaded_file = st.file_uploader("Carrega o ficheiro CSV (separador ';')", type="csv")

if uploaded_file:
    try:
        df = pd.read_csv(uploaded_file, sep=';', parse_dates=['Ano e hora'])
    except Exception as e:
        st.error(f"Erro a ler o CSV: {e}")
    else:
        df['Ano Letivo'] = df['Ano e hora'].apply(definir_ano_letivo)

        resumo = (
            df.groupby(['Ano Letivo', 'Atividade', 'Turma', 'Disciplina'])
            .size()
            .reset_index(name='Contagem')
        )

        wb = Workbook()

        ws_dados = wb.active
        ws_dados.title = "DadosTratados"
        for r in dataframe_to_rows(df, index=False, header=True):
            ws_dados.append(r)

        ws_resumo = wb.create_sheet("Resumo")
        for r in dataframe_to_rows(resumo, index=False, header=True):
            ws_resumo.append(r)

        excel_io = BytesIO()
        wb.save(excel_io)
        excel_io.seek(0)

        st.download_button(
            label="Descarregar Excel com contagens",
            data=excel_io,
            file_name="Contagem_Atividades_Professores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
