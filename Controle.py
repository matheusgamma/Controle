import pandas as pd
import streamlit as st
import os

# Configuração da interface do Streamlit
st.title("Comparador de Planilhas")

# Input para o caminho do arquivo Base Gamma
caminho_base_gamma = st.text_input("Caminho do arquivo Base Gamma:")
# Input para o caminho do arquivo Positivador Novo
caminho_positivador_novo = st.text_input("Caminho do arquivo Positivador Novo:")
# Input para o caminho do arquivo Inclusões
caminho_inclusoes = st.text_input("Caminho do arquivo Inclusões:")

if st.button("Comparar"):
    try:
        # Remover espaços em branco e verificar se o caminho é válido
        caminho_base_gamma = caminho_base_gamma.strip().strip('"')
        caminho_positivador_novo = caminho_positivador_novo.strip().strip('"')
        caminho_inclusoes = caminho_inclusoes.strip().strip('"')

        # Verificar se os arquivos existem
        if not os.path.isfile(caminho_base_gamma):
            st.error(f"O arquivo Base Gamma não foi encontrado: {caminho_base_gamma}")
        elif not os.path.isfile(caminho_positivador_novo):
            st.error(f"O arquivo Positivador Novo não foi encontrado: {caminho_positivador_novo}")
        elif not os.path.isfile(caminho_inclusoes):
            st.error(f"O arquivo Inclusões não foi encontrado: {caminho_inclusoes}")
        else:
            # Ler os arquivos Excel, especificando a aba correta da Base Gamma
            base_gamma = pd.read_excel(caminho_base_gamma, sheet_name="Clientes.Responsáveis")
            positivador_novo = pd.read_excel(caminho_positivador_novo)
            inclusoes_novo = pd.read_excel(caminho_inclusoes)

            # Função para ajustar o código do cliente
            def ajustar_codigo_cliente(codigo):
                if pd.isna(codigo):  # Verifica se é NaN
                    return None  # Retorna None se for NaN
                if isinstance(codigo, float):
                    return str(int(codigo))  # Remove o ".0" ao converter float para int e depois para string
                return str(codigo).strip()  # Garante que qualquer outro código seja tratado como string

            # Ajustar as colunas de clientes para remoção do ".0" e garantir que sejam strings
            base_gamma['Cliente'] = base_gamma['Cliente'].astype(str).str.strip().str.lower()
            positivador_novo['Cliente'] = positivador_novo['Cliente'].astype(str).str.strip().str.lower()
            inclusoes_novo['CODIGO DO CLIENTE'] = inclusoes_novo['CODIGO DO CLIENTE'].apply(ajustar_codigo_cliente).str.lower()

            # Remover entradas None (códigos que eram NaN)
            inclusoes_novo = inclusoes_novo[inclusoes_novo['CODIGO DO CLIENTE'].notna()]

            # Tratar colunas com datetime para evitar erros na conversão
            if 'DATA DO PROCESSO' in base_gamma.columns:
                base_gamma['DATA DO PROCESSO'] = pd.to_datetime(base_gamma['DATA DO PROCESSO'], errors='coerce')

            # Verificar o número de colunas e ajustar a renomeação da Base Gamma
            if len(base_gamma.columns) == 8:
                base_gamma.columns = ['Positivador', 'Dt Entrada', 'Dt Saída', 'Cliente', 'Classe', 'Nome', 'Farmer / Hunter', 'Trader']
            else:
                st.warning(f"Base Gamma possui {len(base_gamma.columns)} colunas. Verifique a estrutura dos dados: {base_gamma.columns.tolist()}")

            # Interseção de clientes
            clientes_base_gamma = set(base_gamma['Cliente'])
            clientes_positivador_novo = set(positivador_novo['Cliente'])
            clientes_inclusoes = set(inclusoes_novo['CODIGO DO CLIENTE'])

            # Interseções corretas entre os grupos
            coincidentes = clientes_base_gamma.intersection(clientes_positivador_novo)
            novos = clientes_positivador_novo - clientes_base_gamma
            saidas = clientes_base_gamma - clientes_positivador_novo
            coincidentes_inclusoes_novos = clientes_inclusoes.intersection(novos)

            # Criar um resumo com os totais
            total_coincidentes = len(coincidentes)
            total_novos = len(novos)
            total_saidas = len(saidas)
            total_inclusoes = len(clientes_inclusoes)
            total_basegamma = len(clientes_base_gamma)
            total_coincidentes_inclusoes_novos = len(coincidentes_inclusoes_novos)

            # Exibir o resumo dos totais
            st.subheader("Resumo dos Totais")
            st.write(f"Total de clientes coincidentes: {total_coincidentes}")
            st.write(f"Total de novos clientes: {total_novos}")
            st.write(f"Total de clientes que saíram: {total_saidas}")
            st.write(f"Total de inclusões: {total_inclusoes}")
            st.write(f"Total Base Gamma: {total_basegamma}")
            st.write(f"Total de clientes coincidentes entre Inclusões e Novos Clientes: {total_coincidentes_inclusoes_novos}")

            # Exibir os DataFrames completos (opcional)
            st.subheader("Coincidentes entre Inclusões e Novos Clientes:")
            st.dataframe(inclusoes_novo[inclusoes_novo['CODIGO DO CLIENTE'].isin(coincidentes_inclusoes_novos)])

            st.subheader("Detalhes - Coincidentes:")
            st.dataframe(base_gamma[base_gamma['Cliente'].isin(coincidentes)])

            st.subheader("Detalhes - Novos Clientes:")
            st.dataframe(positivador_novo[positivador_novo['Cliente'].isin(novos)])

            st.subheader("Detalhes - Clientes que Saíram:")
            st.dataframe(base_gamma[base_gamma['Cliente'].isin(saidas)])

            st.subheader("Detalhes - Inclusões:")
            st.dataframe(inclusoes_novo[inclusoes_novo['CODIGO DO CLIENTE'].isin(clientes_inclusoes)])

            st.success("Comparação e resumo dos totais concluídos com sucesso!")

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
