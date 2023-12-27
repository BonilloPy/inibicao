# -*- coding: utf-8 -*-

import re
import time
import os
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO

from pathlib import Path
import pandas as pd
import streamlit as st

def convert_csv_to_xlsx(folder_path):
    folder = Path(folder_path)
    for file in folder.glob('*.csv'):
        for encoding in ['utf-8', 'latin1', 'ISO-8859-1', 'cp1252']:
            try:
                df = pd.read_csv(file, encoding=encoding, delimiter=',', on_bad_lines='skip')
                xlsx_file = file.with_suffix('.xlsx')
                df.to_excel(xlsx_file, index=False)
                st.success(f"{file.name} convertido para o formato XLSX usando {encoding} encoding.")
                break
            except UnicodeDecodeError:
                continue
            except pd.errors.ParserError:
                st.error(f"Erro ao analisar {file.name}.")
                break
            except PermissionError:
                st.error(f"Erro de permissão ao acessar {file.name}.")
                break
            except FileNotFoundError:
                st.error(f"Arquivo {file.name} não encontrado.")
                break
            except Exception as e:
                st.error(f"Erro ao processar {file.name}: {e}")
                break
        else:
            st.error(f"Não foi possível converter {file.name}. Encoding não suportado.")

### Botão para converter arquivos CSV para XLSX
##if st.button("Converter arquivos CSV para XLSX"):
##    folder_path = st.text_input("Digite o caminho da pasta contendo arquivos CSV")
##    if folder_path:
##        convert_csv_to_xlsx(folder_path)
##        st.write(f"Verificando o diretório: {folder_path}")
##    else:
##        st.error("Por favor, insira um caminho de pasta válido.")




# Define as funções necessárias para processamento de dados
def remove_espacos_colunas(dataframe):
    dataframe.columns = dataframe.columns.str.strip()
    return dataframe

def remove_espacos_celulas(dataframe):
    for col in dataframe.columns:
        if dataframe[col].dtype == 'object':  # Verifica se a coluna é do tipo 'object'
            try:
                dataframe[col] = dataframe[col].str.strip()  # Aplica strip se for string
            except AttributeError:
                # Se não for string, ignora a operação
                pass
    return dataframe


def remove_numeros(texto):
    return re.sub(r'\d+', '', texto) if isinstance(texto, str) else texto

def converter(texto):
    return datetime.strptime(texto, '%b %d %Y %I:%M%p')

def extrair_data(texto):
    if len(texto) == 10:
        return texto[:10].strip()
    else:
        return texto[:11].strip()

def extrair_serie(texto):
    texto = str(texto)
    return texto[:14]

def converte_data_hora(texto):
    if pd.isna(texto) or texto == 'nan':
        return pd.NaT
    texto = " ".join(texto.split())
    lista = texto.split(" ")[:3]
    resultado = ' '.join(map(str, lista))
    try:
        resultado_final = datetime.strptime(resultado, '%b %d %Y')
    except ValueError:
        return pd.NaT
    return resultado_final

def definir_prioridade(row):
    if row['deNivelTecnico'] == 'Técnico Revenda' and row['deTipContrato'] == 'Locação Orgãos Públicos' and row['TOB'] == 'SIM':
        return 1
    elif row['deNivelTecnico'] == 'Técnico Revenda' and row['deTipContrato'] == 'Locação Orgãos Públicos':
        return 2
    elif row['deNivelTecnico'] == 'Técnico Revenda' and row['deTipContrato'] != 'Locação Orgãos Públicos' and row['TOB'] == 'SIM':
        return 3
    elif row['deNivelTecnico'] == 'Técnico Revenda' and row['deTipContrato'] != 'Locação Orgãos Públicos':
        return 4
    elif row['deNivelTecnico'] != 'Técnico Revenda' and row['deTipContrato'] == 'Locação Orgãos Públicos' and row['TOB'] == 'SIM':
        return 5
    elif row['deNivelTecnico'] != 'Técnico Revenda' and row['deTipContrato'] == 'Locação Orgãos Públicos':
        return 6
    elif row['TOB'] == 'SIM':
        return 7
    else:
        return 8

# Configuração do Streamlit
st.title("CONTROLE DE MP")

# Upload dos arquivos via Streamlit
file1 = st.file_uploader("Selecione o arquivo de Relatório de MP's:", type=["xlsx"])
file2 = st.file_uploader("Selecione o segundo arquivo com a LISTA DE PEÇAS:", type=["xlsx"])
file3 = st.file_uploader("Selecione o arquivo do SDS:", type=["xlsx"])
file4 = st.file_uploader("Selecione o arquivo do NDD:", type=["xlsx"])
file5 = st.file_uploader("Selecione o arquivo coma  lista de contratos TOB :", type=["xlsx"])

progress_bar = st.progress(0)

# Botão para iniciar o processo
if st.button('Clique aqui para executar o processo'):
    if file1 and file2 and file3 and file4 and file5:
        with st.spinner('Processando...'):
            start = time.time()

            df1 = pd.read_excel(file1)
            df2 = pd.read_excel(file2)
            df3 = pd.read_excel(file3)
            df4 = pd.read_excel(file4)
            df5 = pd.read_excel(file5, sheet_name='LISTA_CONTRATOS')

            # Processamento dos DataFrames
            df1 = remove_espacos_colunas(df1)
            df2 = remove_espacos_colunas(df2)
            df3 = remove_espacos_colunas(df3)
            df4 = remove_espacos_colunas(df4)

            df1 = remove_espacos_celulas(df1)
            df2 = remove_espacos_celulas(df2)
            df3 = remove_espacos_celulas(df3)
            df4 = remove_espacos_celulas(df4)

            df1.rename(columns={"Nº de Serie": "Serial Number"}, inplace=True)
            df1['Técnico'] = df1['Técnico'].apply(remove_numeros)

            if 'PECAS_NOMES' in df2.columns and 'Descrição do Produto' in df1.columns:
                df1 = df1[df1['Descrição do Produto'].isin(df2['PECAS_NOMES'])]

            df1['Data da Próxima MP'] = pd.to_datetime(df1['Data da Próxima MP'], errors='coerce')
            df1['Data Ultima Leitura'] = pd.to_datetime(df1['Data Ultima Leitura'], errors='coerce')
            df1['dtFimVigencia'] = pd.to_datetime(df1['dtFimVigencia'], errors='coerce')

            conjunto_contratos = set(df5['CONTRATO'].values)
            df1['TOB'] = df1['Contrato'].apply(lambda x: "SIM" if x in conjunto_contratos else "NÃO")

            df1['PRIORIDADE'] = df1.apply(definir_prioridade, axis=1)

            # Preparação e download do arquivo processado
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df1.to_excel(writer, index=False)
            progress_data = output.getvalue()

            progress_bar.progress(100)
            end = time.time()
            decorrido = (end - start) / 60
            st.success(f"Processamento concluído. Tempo decorrido: {decorrido:.2f} minutos.")

            st.download_button(label="Baixar arquivo processado",
                               data=progress_data,
                               file_name='dados_processados.xlsx',
                               mime='application/vnd.ms-excel')

# O resto do seu código original permanece o mesmo.
