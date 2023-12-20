import streamlit as st
import os
import pandas as pd

# Função para converter todos os arquivos CSV em uma pasta para o formato XLSX
def convert_all_csv_to_xlsx(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith('.csv'):
            csv_file_path = os.path.join(folder_path, filename)
            xlsx_file_path = os.path.join(folder_path, filename.replace('.csv', '.xlsx'))

            separator = ';' if 'MP' in filename else ','

            try:
                df = pd.read_csv(csv_file_path, sep=separator)
            except UnicodeDecodeError:
                try:
                    df = pd.read_csv(csv_file_path, sep=separator, encoding='ISO-8859-1')
                except Exception as e:
                    st.error(f"Erro ao processar o arquivo {filename} com a codificação 'ISO-8859-1': {e}")
                    continue

            df.to_excel(xlsx_file_path, index=False)

# Interface Streamlit
st.title('Conversor de CSV para XLSX')

# Campo para inserir o caminho da pasta
folder_path = st.text_input('Informe o caminho da pasta contendo os arquivos CSV:')

# Botão para executar a conversão
if st.button('Converter CSV para XLSX'):
    if os.path.isdir(folder_path):
        with st.spinner('Convertendo arquivos. Por favor, aguarde...'):
            convert_all_csv_to_xlsx(folder_path)
        st.success('Conversão concluída com sucesso!')
    else:
        st.error('O caminho informado não é uma pasta válida.')
