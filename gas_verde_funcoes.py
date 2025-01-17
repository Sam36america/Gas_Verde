from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import pandas as pd
import numpy as np
import os
import re
from sqlalchemy import create_engine, select, MetaData, Table
from sqlalchemy.orm import sessionmaker
import shutil

#in/out 3100, 1300, 3800, 1450
usuario_conectado = 'samuel.santos'
# Configure o caminho do executável do Tesseract
pytesseract.pytesseract.tesseract_cmd = fr'C:\Users\{usuario_conectado}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

def pdf_ocr(image):
    # Select the first page
    config = pytesseract.pytesseract.tesseract_cmd
    
    # Define o idioma para o reconhecimento de texto (por exemplo, português)
    config += r'--oem 3 --psm 6 -l por'
    config += r'--psm 6 outputbase alphanumeric'

    return pytesseract.image_to_string(image,config=config)

def pdf_to_image(pdf_path):

    # Converte o PDF em uma lista de imagens
    images = convert_from_path(pdf_path, 500, poppler_path=r'C:\poppler-0.68.0\bin')
    imagem = images[0]
    #imagem.show()
    return imagem
      
def dados_excel(cnpj, valor_total,volume_total, data_emissao, data_inicio, data_fim, numero_fatura, valor_icms, correcao_pcs, dist):
   
    dados = {
           'CNPJ': [cnpj],
           'VALOR TOTAL': valor_total,
           'VOLUME TOTAL': [volume_total],
           'DATA DA EMISSÃO': data_emissao,
           'DATA INICIO': [data_inicio],
           'DATA FIM': [data_fim],     
           'NUMERO FATURA':[numero_fatura],
           'VALOR ICMS': [valor_icms],
           'CORREÇÃO DO PCS': [correcao_pcs],
           'DISTRIBUIDORA': [dist]
     }
    try:    
        df = pd.DataFrame(dados)
    except:
        dados = {
           'CNPJ':'CNPJ não encontrado', 
           'VALOR TOTAL': 'valor_total não econtrado',
           'VOLUME TOTAL': 'volume_total não econtrado',
           'DATA DA EMISSÃO': 'data_emissao não econtrado',
           'DATA INICIO': 'data_inicio não econtrado',
           'DATA FIM': 'data_fim não econtrado',     
           'NUMERO FATURA':'numero_fatura não econtrado]',
           'VALOR ICMS': 'valor_icms não econtrado',
           'CORREÇÃO PCS':'Correção pcs não encontrado',
           'DISTRIBUIDORA': 'distribuidora não econtrado'
            }
        
        indice = ['1']
        df = pd.DataFrame(dados,index=indice)  
    
    return df
          
def adicionar_dados_excel(dados, novos_dados):
    try:
        df_existente = pd.read_excel(dados)
    except FileNotFoundError:
        print(f"O arquivo '{dados}' não foi encontrado. Criando um novo.")
        df_existente = pd.DataFrame()
    
    try:
        df_novos_dados = pd.DataFrame(novos_dados)
        
        # Converte as colunas específicas para numérico
        colunas_numericas = ['VALOR TOTAL', 'VOLUME TOTAL', 'VALOR ICMS']
        for coluna in colunas_numericas:
            if coluna in df_novos_dados.columns:
                df_novos_dados[coluna] = pd.to_numeric(df_novos_dados[coluna].str.replace('.', '').str.replace(',', '.'))
        
        df_resultante = pd.concat([df_existente, df_novos_dados], ignore_index=True)
        df_resultante.to_excel(dados, index=False)
        print(f"Dados adicionados com sucesso na planilha '{dados}'")
        return True
    except Exception as e:
        print(f"Erro ao adicionar os dados na planilha '{dados}': {e}")
        return False

def listar_pdfs_com_referencia_na_pasta(pasta, referencia):
    arquivos_pdf = []
    for arquivo in os.listdir(pasta):
        if arquivo.endswith('.pdf'):
            nome_distribuidora = re.findall(r'_GN_([A-ZÁ]+)_',arquivo)
            if nome_distribuidora:
                nome_distribuidora = nome_distribuidora[0]
                
            arquivos_pdf.append(arquivo)
    return arquivos_pdf

def verificar_fatura_existe(session, tabela_faturas, numero_fatura):
    stmt = select([tabela_faturas.c.numero_fatura]).where(tabela_faturas.c.numero_fatura == numero_fatura)
    result = session.execute(stmt).fetchone()
    return result is not None

def data_fim_mes(data_fim):
    meses = {
        'JANEIRO' : '31/01',
        'FEVEREIRO': '29/02',  # Considerando ano bissexto
        'MARÇO': '31/03',
        'MARCO': '31/03',
        'ABRIL': '30/04',
        'MAIO': '31/05',
        'JUNHO': '30/06',
        'JULHO': '31/07',
        'AGOSTO': '31/08',
        'SETEMBRO': '30/09',
        'OUTUBRO': '31/10',
        'NOVEMBRO': '30/11',
        'DEZEMBRO': '31/12'
    }
    mes, ano = data_fim.split('/')
    data_fim = f'{meses[mes]}/{ano}'
    return data_fim

def data_inicio_mes(data_fim):
    meses = {
        'JANEIRO': '01/01',
        'FEVEREIRO': '01/02',
        'MARÇO': '01/03',
        'MARCO': '01/03',
        'ABRIL': '01/04',
        'MAIO': '01/05',
        'JUNHO': '01/06',
        'JULHO': '01/07',
        'AGOSTO': '01/08',
        'SETEMBRO': '01/09',
        'OUTUBRO': '01/10',
        'NOVEMBRO': '01/11',
        'DEZEMBRO': '01/12'
    }
    mes, ano = data_fim.split('/')
    data_inicio = f'{meses[mes]}/{ano}'
    return data_inicio

def verificar_download(cnpj, data_inicio, data_fim, excel_path):
    # Carregar o arquivo Excel
    df = pd.read_excel(excel_path, sheet_name='Sheet1')
    
    cnpj = int(cnpj)

    # Filtrar as linhas que correspondem aos critérios
    df_filtrado = df[
        (df['CNPJ'] == cnpj) &
        (df['DATA INICIO'] == data_inicio) &
        (df['DATA FIM'] == data_fim)
    ]
    
    # Verificar se há pelo menos uma linha que atenda aos critérios
    if len(df_filtrado) > 0:
        return False
    else:
        return True

def mover_faturas_lidas(file_path, diretorio_destino):
    if os.path.isfile(file_path):
        try:
            # Move o arquivo para o diretório de destino
            shutil.move(file_path, diretorio_destino)
            print(f'Arquivo movido para {file_path}')
        except Exception as e:
            print(f'Erro ao mover o arquivo para {file_path}: {e}')

# Caminho da pasta onde os PDFs estão localizados
pasta_pdfs = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Gás Verde\Faturas'

# Referência que queremos encontrar nos nomes dos arquivos
referencia = 'GÁS VERDE'  # Exemplo: todohnbs os arquivos que contêm 'FATURA' no nome

# Listar todos os PDFs na pasta que contêm a referência no nome
pdfs_com_referencia = listar_pdfs_com_referencia_na_pasta(pasta_pdfs, referencia)

# Exibir os PDFs encontrados
for pdf in pdfs_com_referencia:
    print(pdf)



