import pandas as pd
import streamlit as st
import openpyxl

# Configuração da interface do Streamlit
st.title("Relatório de movimentos - Gamma Capital")

# Input para o arquivo Base Gamma
arquivo_base_gamma = st.file_uploader("Carregue o arquivo Base Gamma:", type=["xlsx"])
# Input para o arquivo Positivador Novo
arquivo_positivador_novo = st.file_uploader("Carregue o arquivo Positivador Novo:", type=["xlsx"])
# Input para o arquivo Inclusões
arquivo_inclusoes = st.file_uploader("Carregue o arquivo Inclusões:", type=["xlsx"])

if st.button("Comparar"):
    try:
        # Verificar se os arquivos foram carregados
        if arquivo_base_gamma is None:
            st.error("Por favor, carregue o arquivo Base Gamma.")
        elif arquivo_positivador_novo is None:
            st.error("Por favor, carregue o arquivo Positivador Novo.")
        elif arquivo_inclusoes is None:
            st.error("Por favor, carregue o arquivo Inclusões.")
        else:
            # Ler os arquivos Excel diretamente do UploadedFile
            base_gamma = pd.read_excel(arquivo_base_gamma, sheet_name="Master")
            positivador_novo = pd.read_excel(arquivo_positivador_novo)
            inclusoes_novo = pd.read_excel(arquivo_inclusoes)


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
            inclusoes_novo['CODIGO DO CLIENTE'] = inclusoes_novo['CODIGO DO CLIENTE'].apply(
                ajustar_codigo_cliente).str.lower()

            # Remover entradas None (códigos que eram NaN)
            inclusoes_novo = inclusoes_novo[inclusoes_novo['CODIGO DO CLIENTE'].notna()]

            # Verificar o número de colunas e ajustar a renomeação da Base Gamma
            if len(base_gamma.columns) == 8:
                base_gamma.columns = ['Positivador', 'Dt Entrada', 'Dt Saída', 'Cliente', 'Classe', 'Nome',
                                      'Farmer / Hunter', 'Trader']
            else:
                st.warning(
                    f"Base Gamma possui {len(base_gamma.columns)} colunas. Verifique a estrutura dos dados: {base_gamma.columns.tolist()}")

            # Remover duplicatas na Base Gamma
            base_gamma = base_gamma.drop_duplicates(subset='Cliente')

            # Interseção de clientes
            clientes_base_gamma = set(base_gamma['Cliente'])
            clientes_positivador_novo = set(positivador_novo['Cliente'])
            clientes_inclusoes = set(inclusoes_novo['CODIGO DO CLIENTE'])

            # Interseções corretas entre os grupos
            coincidentes = clientes_base_gamma.intersection(clientes_positivador_novo)
            novos = clientes_positivador_novo - clientes_base_gamma
            saidas = clientes_base_gamma - clientes_positivador_novo
            coincidentes_inclusoes_novos = clientes_inclusoes.intersection(novos)

            # Cálculo do relatório de movimentações
            inicio_total = len(clientes_base_gamma)
            entradas_total = len(novos)
            saidas_total = len(saidas)
            fim_total = inicio_total + entradas_total - saidas_total

            dentro_positivador_inicio = len(coincidentes)
            dentro_positivador_saidas = len(saidas.intersection(coincidentes))
            dentro_positivador_fim = dentro_positivador_inicio - dentro_positivador_saidas

            fora_positivador_inicio = inicio_total - dentro_positivador_inicio
            fora_positivador_saidas = len(saidas.difference(coincidentes))
            fora_positivador_entradas = len(novos)
            fora_positivador_fim = fora_positivador_inicio + fora_positivador_entradas - fora_positivador_saidas

            # Criar DataFrame para o Relatório de Movimentações
            relatorio = pd.DataFrame({
                '': ['Total', 'Dentro Positivador', 'Fora Positivador'],
                'Início': [inicio_total, dentro_positivador_inicio, fora_positivador_inicio],
                'Entradas': [entradas_total, 0, fora_positivador_entradas],
                'Saídas': [saidas_total, dentro_positivador_saidas, fora_positivador_saidas],
                'Fim': [fim_total, dentro_positivador_fim, fora_positivador_fim]
            })

            # Calcular informações extras (Entradas, Coincidente Inclusões, Sem Farmer)
            total_coincidentes_inclusoes_novos = len(coincidentes_inclusoes_novos)
            sem_farmer = entradas_total - total_coincidentes_inclusoes_novos

            # Criar DataFrame para as novas informações no topo
            informacoes_extras = pd.DataFrame({
                '': ['Entradas', 'Coincidente Inclusões', 'Sem Farmer'],
                'Total': [entradas_total, total_coincidentes_inclusoes_novos, sem_farmer]
            })

            # --- Painel 1: Resumo e Movimentações ---
            with st.expander("Painel 1: Base Gamma x Positivador", expanded=True):
                st.subheader("Resumo dos Totais")
                st.write(f"Total de clientes coincidentes (Positivador x Base Gamma): {len(coincidentes)}")
                st.write(f"Total de novos clientes: {entradas_total}")
                st.write(f"Total de clientes que saíram: {len(saidas)}")
                st.write(f"Total de inclusões: {len(clientes_inclusoes)}")
                st.write(f"Total Base Gamma: {inicio_total}")
                st.write(
                    f"Total de clientes coincidentes entre Inclusões e Novos Clientes: {total_coincidentes_inclusoes_novos}")

                st.subheader("Coincidentes Positivador e Base Gamma:")
                st.dataframe(base_gamma[base_gamma['Cliente'].isin(coincidentes)])

                st.subheader("Detalhes - Novos Clientes:")
                st.dataframe(positivador_novo[positivador_novo['Cliente'].isin(novos)])

                st.subheader("Detalhes - Clientes que Saíram:")
                st.dataframe(base_gamma[base_gamma['Cliente'].isin(saidas)])

                st.subheader("Relatório de Movimentações")
                st.dataframe(relatorio)

                st.subheader("Informações Adicionais")
                st.dataframe(informacoes_extras)

            # --- Painel 2: Coincidências ---
            with st.expander("Painel 2: Coincidências", expanded=False):

                st.subheader("Coincidentes entre Inclusões e Novos Clientes:")
                st.dataframe(inclusoes_novo[inclusoes_novo['CODIGO DO CLIENTE'].isin(coincidentes_inclusoes_novos)])

                # Calcular diferença entre novos e coincidências
                diferenca_novos_coincidentes = novos - coincidentes_inclusoes_novos
                st.subheader("Clientes sem farmer:")
                st.dataframe(positivador_novo[positivador_novo['Cliente'].isin(diferenca_novos_coincidentes)])


            # --- Painel 3: Outros Detalhes ---
            with st.expander("Painel 3: Outros Detalhes", expanded=False):

                st.subheader("Detalhes - Inclusões:")
                st.dataframe(inclusoes_novo[inclusoes_novo['CODIGO DO CLIENTE'].isin(clientes_inclusoes)])

            # Agrupar clientes por Classe e Farmer/Hunter
            clientes_por_pessoa = base_gamma.groupby(['Classe', 'Farmer / Hunter']).size().reset_index(
                name='Total de Clientes')

            # Somar todos os clientes por Classe
            total_geral = clientes_por_pessoa['Total de Clientes'].sum()

            # Adicionar uma linha de total geral
            total_row = pd.DataFrame([['Total', '', total_geral]],
                                     columns=['Classe', 'Farmer / Hunter', 'Total de Clientes'])

            clientes_por_pessoa = pd.concat([clientes_por_pessoa, total_row], ignore_index=True)

            # --- Painel 4: Tabela dinâmica (Clientes por Classe e Pessoa) ---
            st.subheader("Painel 4: Clientes por Classe e Pessoa")

            # Loop para criar expanders para cada classe
            classes = clientes_por_pessoa['Classe'].unique()

            for classe in classes:
                with st.expander(f"{classe}", expanded=False):
                    # Filtrar clientes por classe específica
                    clientes_classe = clientes_por_pessoa[clientes_por_pessoa['Classe'] == classe]

                    # Calcular subtotal da classe
                    subtotal = clientes_classe['Total de Clientes'].sum()

                    # Exibir os dados da classe expandida
                    st.dataframe(clientes_classe)

                    # Exibir o subtotal da classe
                    st.write(f"Subtotal para {classe}: {subtotal} clientes")


            st.success("Comparação e relatório de movimentações gerados com sucesso!")

    except Exception as e:
        st.error(f"Ocorreu um erro durante a comparação: {e}")
