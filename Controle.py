import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime
from io import BytesIO

# Configuração do Streamlit
st.set_page_config(page_title="Relatório de Transferências de Clientes", layout="wide", page_icon="📊")

# Título
st.title("📊 Relatório Consolidado de Transferências de Clientes por Assessor")

# Upload dos arquivos
arquivo_master = st.file_uploader("Carregue a Master:", type=["xlsm"])
arquivo_transferencia = st.file_uploader("Carregue o arquivo de Transferências de Contas:", type=["xlsx"])

# Seleção do mês/ano
mes_ano = st.text_input("Digite o mês/ano no formato MM/AAAA:", value="11/2024")

with st.spinner('Comparando as bases de dados, aguarde...'):
    if st.button("Gerar Relatório"):
        if None in [arquivo_master, arquivo_transferencia]:
            st.error("Por favor, carregue os dois arquivos.")
        else:
            # Carregar os arquivos
            master = pd.read_excel(arquivo_master, sheet_name="Clientes.Responsáveis")
            transferencia = pd.read_excel(arquivo_transferencia)

            # Filtrar transferências concluídas
            transferencia = transferencia[transferencia['Status'] == 'CONCLUIDO']

            master = master.rename(columns={'Cliente': 'Código do Cliente'})

            # Vincular clientes ao Assessor na Master
            clientes_master = master[['Código do Cliente', 'Farmer']].drop_duplicates()
            transferencia = transferencia.merge(clientes_master, on='Código do Cliente', how='left', suffixes=('', '_Master'))

            # Preencher "Não Consta" para clientes que não foram encontrados na Master
            transferencia['Farmer'] = transferencia['Farmer'].fillna("Não Consta")

            # Filtrar pelo mês/ano
            transferencia['Data Transferência'] = pd.to_datetime(transferencia['Data Transferência'], format="%d/%m/%Y %H:%M:%S")
            transferencia['Mês/Ano'] = transferencia['Data Transferência'].dt.strftime("%m/%Y")
            transferencia_filtrada = transferencia[transferencia['Mês/Ano'] == mes_ano]

            # Identificar entradas e saídas
            entradas = transferencia_filtrada[
                transferencia_filtrada['Nome Assessor Origem'].isna() &
                transferencia_filtrada['Nome Assessor Destino'].notna()
            ]
            saidas = transferencia_filtrada[
                transferencia_filtrada['Nome Assessor Origem'].notna() &
                transferencia_filtrada['Nome Assessor Destino'].isna()
            ]

            # Contar clientes por assessor
            clientes_por_assessor = master.groupby('Farmer').size().reset_index(name="Clientes Iniciais")
            entradas_por_assessor = entradas.groupby('Farmer')['Código do Cliente'].nunique().reset_index(name="Entradas")
            saidas_por_assessor = saidas.groupby('Farmer')['Código do Cliente'].nunique().reset_index(name="Saídas")

            # Consolidar resumo por assessor
            todos_assessores = list(master['Farmer'].unique()) + ["Não Consta"]
            resumo_por_assessor = pd.DataFrame({'Farmer': todos_assessores}).merge(
                clientes_por_assessor,
                on="Farmer", how="left"
            ).merge(
                entradas_por_assessor,
                on="Farmer", how="left"
            ).merge(
                saidas_por_assessor,
                on="Farmer", how="left"
            )

            # Ajustar valores para zero onde necessário
            resumo_por_assessor = resumo_por_assessor.fillna(0)
            resumo_por_assessor['Entradas'] = resumo_por_assessor['Entradas'].astype(int)
            resumo_por_assessor['Saídas'] = resumo_por_assessor['Saídas'].astype(int)
            resumo_por_assessor['Final'] = resumo_por_assessor['Clientes Iniciais'] + resumo_por_assessor['Entradas'] - resumo_por_assessor['Saídas']

            # Adicionar linha de total consolidado
            total_resumo = pd.DataFrame({
                "Farmer": ["Total Consolidado"],
                "Clientes Iniciais": [resumo_por_assessor["Clientes Iniciais"].sum()],
                "Entradas": [resumo_por_assessor["Entradas"].sum()],
                "Saídas": [resumo_por_assessor["Saídas"].sum()],
                "Final": [resumo_por_assessor["Final"].sum()]
            })
            resumo_por_assessor = pd.concat([resumo_por_assessor, total_resumo], ignore_index=True)

            # Criar relatório analítico de entradas e saídas
            analitico_entradas = entradas[['Código do Cliente', 'Farmer', 'Nome Assessor Destino']].rename(
                columns={'Farmer': 'Assessor Origem', 'Nome Assessor Destino': 'Assessor Destino'}
            ).assign(Status="Entrada")
            analitico_saidas = saidas[['Código do Cliente', 'Farmer', 'Nome Assessor Origem']].rename(
                columns={'Farmer': 'Assessor Destino', 'Nome Assessor Origem': 'Assessor Origem'}
            ).assign(Status="Saída")

            detalhes_clientes = pd.concat([analitico_entradas, analitico_saidas], ignore_index=True)
            detalhes_clientes = detalhes_clientes.fillna("Não Consta")

            # Exibir Resumo e Gráficos
            st.subheader("📋 Resumo Consolidado")
            st.metric("Total de Clientes Inicial", resumo_por_assessor["Clientes Iniciais"].iloc[-1])
            st.metric("Total de Entradas", resumo_por_assessor["Entradas"].iloc[-1])
            st.metric("Total de Saídas", resumo_por_assessor["Saídas"].iloc[-1])
            st.metric("Total Final de Clientes", resumo_por_assessor["Final"].iloc[-1])

            st.subheader("📋 Resumo por Assessor")
            st.dataframe(resumo_por_assessor.rename(columns={"Farmer": "Assessor"}))

            # Exportar Relatórios para Excel
            with st.expander("⬇️ Exportar Relatórios"):
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    resumo_por_assessor.to_excel(writer, index=False, sheet_name="Resumo por Assessor")
                    detalhes_clientes.to_excel(writer, index=False, sheet_name="Detalhes Clientes")
                buffer.seek(0)
                st.download_button(
                    label="Baixar Relatórios Consolidado",
                    data=buffer,
                    file_name="relatorios_transferencias.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
