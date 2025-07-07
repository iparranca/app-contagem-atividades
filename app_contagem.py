import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

st.set_page_config(page_title="Contagem de Atividades", layout="wide")
st.title("📊 Contagem de Atividades por Parâmetro2222")

uploaded_file = st.file_uploader("Carregar ficheiro CSV (separador ';')", type=["csv"])

def determinar_ano_letivo(data):
    if data.month >= 9:
        return f"{data.year}/{data.year + 1}"
    else:
        return f"{data.year - 1}/{data.year}"

if uploaded_file:
    try:
        df = pd.read_csv(
            uploaded_file,
            sep=';',
            encoding='latin1',
            parse_dates=['Ano e hora']
        )
    except Exception as e:
        st.error(f"Erro ao ler o CSV: {e}")
        st.stop()

    # Limpar e renomear colunas
    df.columns = df.columns.str.strip()
    df.rename(columns={'Ano e hora': 'DataHora'}, inplace=True)

    # Converter para datetime
    df['DataHora'] = pd.to_datetime(df['DataHora'], errors='coerce')

    # Criar coluna Ano Letivo
    df['AnoLetivo'] = df['DataHora'].apply(determinar_ano_letivo)

    # Mostrar dados carregados
    with st.expander("👁️ Visualizar dados carregados"):
        st.dataframe(df)

    # Seleção do tipo de contagem
    tipo_contagem = st.selectbox(
        "Selecionar tipo de contagem:",
        [
            "Por Atividade",
            "Por Turma",
            "Por Atividade e Turma",
            "Por Disciplina",
            "Por Ano Letivo",
            "Por Atividade e Ano Letivo",
            "Por Disciplina e Turma"
        ]
    )

    # Geração da contagem
    if tipo_contagem == "Por Atividade":
        tabela = df.groupby("Atividade").size().reset_index(name="Contagem")
    elif tipo_contagem == "Por Turma":
        tabela = df.groupby("Turma").size().reset_index(name="Contagem")
    elif tipo_contagem == "Por Atividade e Turma":
        tabela = df.groupby(["Atividade", "Turma"]).size().reset_index(name="Contagem")
    elif tipo_contagem == "Por Disciplina":
        tabela = df.groupby("Disciplina").size().reset_index(name="Contagem")
    elif tipo_contagem == "Por Ano Letivo":
        tabela = df.groupby("AnoLetivo").size().reset_index(name="Contagem")
    elif tipo_contagem == "Por Atividade e Ano Letivo":
        tabela = df.groupby(["Atividade", "AnoLetivo"]).size().reset_index(name="Contagem")
    elif tipo_contagem == "Por Disciplina e Turma":
        tabela = df.groupby(["Disciplina", "Turma"]).size().reset_index(name="Contagem")
    else:
        tabela = pd.DataFrame()

    # Mostrar resultado
    st.subheader("📋 Resultado da Contagem")
    st.dataframe(tabela)

    # Exportar para Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        tabela.to_excel(writer, index=False, sheet_name="Contagem")
    output.seek(0)

    st.download_button(
        label="📥 Download da tabela em Excel",
        data=output.read(),
        file_name="contagem_atividades.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👆 Carrega um ficheiro CSV para começar.")
