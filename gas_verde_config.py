# COORDENADAS PARA EXTRAÇÃO DE DADOS USANDO OCR

def corte_gas_verde():
    corte = {

    'cnpj': (2550, 1800, 3400, 2800), #
    'cnpj_ajustado': (2500, 1550, 3530, 1800),

    'valor_total': (3520, 2660, 3950, 2800),#
    'valor_total_ajustado':(),

    'volume_total': (2040, 3440, 2300, 3530), #
    'volume_total_ajustado':  (2040, 2990, 2285, 3100),
    'volume_total_ajustado2': (2000, 2890, 2385, 3200),
     
    'data_emissao': (3330, 1595, 4000, 1800), #

    'data_inicio': (520, 4650, 900, 4700), 
    'data_inicio_ajustado': (505, 4680, 1800, 4830),
    'data_inicio_ajustado2': (100, 4650, 1700, 4750),

    'data_fim': (520, 4650, 900, 4700),
    'data_fim_ajustado': (272, 4680, 1800, 4830),
    'data_fim_ajustado2': (300, 4650, 1700, 4750),

    'numero_fatura': (3500, 200, 4000, 400),
    'numero_fatura_ajustado': (000, 430, 3750, 590), 
    
    'valor_icms': (1000, 2660, 1610, 2800)
    }
    return corte
caminho_excel = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\00 Faturas Lidas\GAS VERDE.xlsx'
'''while True
    for Player(get_moeda):
        if player(moeda >= 3):
            brilhar(dourado)
        else:
            brilhar(branco)'''


