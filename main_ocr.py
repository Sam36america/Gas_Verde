from PIL import Image
import pytesseract
import re
import os
import pandas as pd
import numpy as np
from gas_verde_config import corte_gas_verde, caminho_excel
from gas_verde_funcoes import pdf_ocr, pdf_to_image, dados_excel, adicionar_dados_excel, data_fim_mes, data_inicio_mes, verificar_download, mover_faturas_lidas

DIST = 'GÁS VERDE'
correcao_pcs = ''
#in/out 3100, 1300, 3800, 1450
#X-Y X-Y
usuario_conectado = 'samuel.santos'
# Configure o caminho do executável do Tesseract
pytesseract.pytesseract.tesseract_cmd = fr'C:\Users\{usuario_conectado}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

def extrator_cnpj(imagem, cordenadas):#
        try:
                cnpj = imagem.crop(corte[cordenadas])
                #cnpj.show()
                cnpj = pdf_ocr(cnpj)
                cnpj = cnpj.replace (',','').replace('/','').replace('-','').replace('.', '')
                cnpj = re.findall(r'(\d+)',cnpj)
                cnpj = ''.join(cnpj)
                cnpj = str(cnpj)

                return cnpj
        except:
                return False

def extrator_valor_total(imagem, coordenadas):#
        try:
                valor_total = imagem.crop(corte[coordenadas])
                #valor_total.show()
                valor_total = pdf_ocr(valor_total)
                valor_total = re.findall(r'(\d{1,3}\.?\d{1,3}\.?\s?\,?\d{1,2})',valor_total)        
                valor_total  = valor_total[0].strip()
                valor_total = valor_total.replace('.','').replace(",",".")
                return valor_total
        except:
                return valor_total
    
def extrator_volume_total(imagem, coordenadas):#     #QUANTIDADE
        try:        
                volume_total = imagem.crop(corte[coordenadas])
                #volume_total.show()
                volume_total = pdf_ocr(volume_total)
                volume_total = re.findall(r'\s?(\d+[\.?A-a\,?]\d+)\s?',volume_total)
                volume_total = volume_total[0].strip()
                volume_total = volume_total.replace('.','').replace(",",".")
                volume_total = round(float(volume_total),5)
                volume_total = str(volume_total)
                
                return volume_total
        
        except:
                return False
    
def extrator_data_emissao(imagem, coordenadas):#
        try:
                data_emissao = imagem.crop(corte[coordenadas])
                #data_emissao.show()
                data_emissao = pdf_ocr(data_emissao)
                data_emissao = re.findall(r'\d{2}/\d{2}/\d{4}',data_emissao)               
                data_emissao  = data_emissao[0].strip()                   

                return data_emissao.split()
        except:
                return False
                
def extrator_data_inicio(imagem, coordenadas):  #?  
        try:
                data_inicio = imagem.crop(corte[coordenadas])
                data_inicio.show()
                data_inicio = pdf_ocr(data_inicio)
                data_inicio = re.findall(r'\s([A-Za-z]+\/\d{4})',data_inicio)
                data_inicio = data_inicio[0].upper()
                data_inicio = data_inicio_mes(data_inicio)

                return data_inicio
        except:
                return False

def extrator_data_fim(imagem, coordenadas): #?

        try:
                data_fim = imagem.crop(corte[coordenadas])
                #data_fim.show()
                data_fim = pdf_ocr(data_fim)
                data_fim = re.findall(r'\s([A-Za-z]+\/\d{4})',data_fim)
                data_fim = data_fim[0].upper()
                data_fim = data_fim_mes(data_fim)
                return data_fim
        except:
                return False
        
def extrator_numero_fatura(imagem, coordenadas): #
        
        try:
                numero_fatura = imagem.crop(corte[coordenadas])
                #numero_fatura.show()
                numero_fatura = pdf_ocr(numero_fatura)
                numero_fatura = re.findall(r'\s?(\d+\.?\d+\.?\d+)',numero_fatura)
                numero_fatura  = numero_fatura[0].strip()
                
                return numero_fatura
        except:
                        
                return False

def extrator_icms (imagem, coordenadas):#  
        
        try:
                valor_icms = imagem.crop(corte[coordenadas])
                #valor_icms.show()
                valor_icms = pdf_ocr(valor_icms)
                valor_icms = re.findall(r'\s?(\d+\.?\,?\d+\.?\,?\d+\.?\,\s?\d+)',valor_icms)
                valor_icms  = valor_icms[0].strip()
                
                return valor_icms
        except:
                        
                return False

def main(pdf_file):
    imagem = pdf_to_image(pdf_file)

    cnpj = extrator_cnpj (imagem, 'cnpj')
    if cnpj == False or len(cnpj) != 14:
        cnpj = extrator_cnpj (imagem, 'cnpj_ajustado')
        if cnpj == False or len (cnpj) != 14:
           cnpj = extrator_cnpj (imagem, 'cnpj_ajustado2')               #continuar daqui FDS
    
    valor_total = extrator_valor_total(imagem, 'valor_total')
    if extrator_valor_total == False:
            valor_total = extrator_valor_total(imagem, 'valor_total_justado')
    
    volume_total = extrator_volume_total(imagem, 'volume_total')
    if volume_total == False:
                volume_total = extrator_volume_total(imagem, 'volume_total_ajustado')
                if volume_total == False:
                        volume_total = extrator_volume_total(imagem, 'volume_total_ajustado2' )

    data_emissao = extrator_data_emissao(imagem, 'data_emissao')
    if data_emissao == False:
            data_emissao = extrator_data_emissao(imagem, 'data_emissao_ajustado')
    
    data_inicio = extrator_data_inicio(imagem, 'data_inicio')
    if data_inicio == False:
                data_inicio = extrator_data_inicio(imagem, 'data_inicio_ajustado')               
                if data_inicio == False:
                        data_inicio = extrator_data_inicio(imagem, 'data_inicio_ajustado2')
                         
    data_fim = extrator_data_fim(imagem, 'data_fim')
    if data_fim == False:
                data_fim = extrator_data_fim(imagem, 'data_fim_ajustado')
                if data_fim == False:
                        data_fim = extrator_data_fim(imagem, 'data_fim_ajustado2')

    numero_fatura = extrator_numero_fatura(imagem, 'numero_fatura')
    if numero_fatura == False:
                numero_fatura = extrator_numero_fatura(imagem, 'numero_fatura_ajustado')

    valor_icms = extrator_icms(imagem, 'valor_icms')
    if valor_icms == False:
                valor_icms = extrator_icms(imagem, 'valor_icms_ajustado')
    
    if not cnpj or not valor_total or not volume_total or not data_emissao or not data_inicio or not data_fim or not numero_fatura or not valor_icms:
        print('Fatura não movida devido a dados incompletos.')
    else: 
        mover_faturas_lidas(pdf_file, diretorio_destino)
        verificar = verificar_download(cnpj, data_inicio, data_fim, caminho_excel)
        if verificar:
                data_frame = dados_excel(cnpj, valor_total, volume_total, data_emissao, data_inicio, data_fim, numero_fatura, valor_icms, correcao_pcs, DIST)
                adicionar_dados_excel(caminho_excel, data_frame)
        else:
                print('Dados já inseridos!')
# Exemplo de uso
corte = corte_gas_verde()
file_path = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Gás Verde\Faturas'
diretorio_destino = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Gás Verde\Lidas'

for arquivo in os.listdir(file_path):
        if arquivo.endswith('.pdf'):
                arquivo = rf'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Gás Verde\Faturas\{arquivo}'
        main(arquivo)
#texto_extraido = pdf_to_image(pdf_path)

#print(texto_extraido)
# Exemplo de uso
