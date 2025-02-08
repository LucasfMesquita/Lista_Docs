#   ALTO LÁ!!
#	QUEM SE APROXIMA DOS PORTÕES DO EGITOOO!!!
#	CUIDADO RAPAZ!!
#	A BARRA AQUI É PESADAAAA!!!!
#
#   Aqui voce verá altas gambiarras, o importante é que funciona
#
#   Se você visitou este script adicione +1 na contagem
#
#   Count (1)
#
#Bibliotecas

import pandas as pd
import openpyxl
import os
import sys
import logging
from logging.handlers import RotatingFileHandler

# Configuracao do Logger com RotatingFileHandler

log_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "script.log")

# Configura o handler de rotacao do arquivo de log

handler = RotatingFileHandler(log_file_path, maxBytes=5*1024*1024, backupCount=5) # 5MB por arquivo, até 5 backups

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        handler,
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger()

# Diretorio do executavel

if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

# Caminho para o arquivo Excel na mesma pasta do executável

excel_path = os.path.join(base_dir, "docs.xlsx")

# Função Principal

def main():

# Carregar o arquivo Excel

    try:
        df = pd.read_excel(excel_path)
        logger.info("Arquivo Excel Carregado")
        logger.debug(f"Tipos de dados: \n{df.dtypes}")
    except FileNotFoundError:
        logger.error(f"Arquivo Excel não encontrado em '{excel_path}'")
        sys.exit(1)

# Limpeza de Dados

    logger.info("Limpando Dados...")
    df['Nº Doc'] = df['Nº Doc'].astype('Int64')
    df_subset = df[['Nº Doc', 'Tipo de Documento']]
    logger.debug(f"Tipos de dados do subset: \n{df_subset.dtypes}")
    logger.info("Retirando Colunas...")
    logger.debug(f"Primeiras 5 linhas do subset: \n{df_subset.head(5)}")
    clean_subset = df_subset.dropna()
    logger.info("Removendo Valores NaN...")
    logger.debug(f"Primeiras 5 linhas do subset limpo: \n{clean_subset.head(5)}")
    df_final = clean_subset.iloc[::-1].reset_index(drop=True)
    df_final.index = range(1, len(df_final) + 1)
    logger.info("DataFrame Final")
    logger.debug(f"Primeiras 5 linhas do DataFrame final: \n{df_final.head(5)}")

# Função para formatar as linhas

    def format_row(row):
        index = row.name
        tipo_documento = row['Tipo de Documento']
        numero_doc = row['Nº Doc']
        if tipo_documento == "Outros documentos em PDF":
            return f"{index}. {tipo_documento}, contendo: xxx,xxx,xxx (Documento Nº {numero_doc})"
        elif tipo_documento == "Outros documentos em demais formatos (JPEG, PNG, DWG e ZIP)":
            return f"{index}. {tipo_documento}, contendo: xxx,xxx,xxx (Documento Nº {numero_doc})"
        else:
            return f"{index}. {tipo_documento} (Documento Nº {numero_doc})"

    logger.info("Criando Arquivo .txt ...")
    linhas_formatadas = df_final.apply(format_row, axis=1).tolist()

# Caminho para salvar o arquivo .txt no mesmo diretório do executável

    txt_path = os.path.join(base_dir, "documentos_formatados.txt")

# Escreva as linhas formatadas em um arquivo .txt (substituindo o conteúdo existente)

    try:
        with open(txt_path, 'w', encoding='utf-8') as file:
            for linha in linhas_formatadas:
                file.write(linha + '\n')
        logger.info(f"As linhas formatadas foram salvas em '{txt_path}'")
    except IOError:
        logger.error(f"Erro ao escrever em '{txt_path}'")

# Abre o arquivo .txt com o Bloco de Notas

    try:
        os.startfile(txt_path)
    except Exception as e:
        logger.error(f"Erro ao abrir o arquivo '{txt_path}' com o Bloco de Notas: {e}")

# Indicação de conclusão do processo e opção de reiniciar ou sair

    opcao = input("\nO processo foi completado. Aperte 1 para analizar novamente ou qualquer outra tecla para sair: ")
    if opcao == "1":
        main()  # Reinicia o programa
    else:
        logger.info("Encerrado o Programa")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.critical("Ocorreu um erro crítico", exc_info=True)
        sys.exit(1)