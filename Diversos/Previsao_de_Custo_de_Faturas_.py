import pandas
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Variáveis Globais

# Variáveis de armazenamento

distribuidoras = ['Celesc', 'Cemig', 'Copel', 'CPFL', 'CPFL Piratininga', 'EDP ES', 'EDP SP', 'Elektro', 'Energisa', 'Equatorial']

revisao_tarifaria = {'Celesc':        [datetime(2026, 5, 1), datetime(2027, 5, 1)],   'Cemig':            [datetime(2026, 7, 1), datetime(2027, 7, 1)], 'Copel':          [datetime(2026, 7, 1), datetime(2027, 7, 1)], 
                     'CPFL':          [datetime(2026, 5, 1), datetime(2027, 5, 1)],   'CPFL Piratininga': [datetime(2026, 5, 1), datetime(2027, 5, 1)], 'EDP ES':         [datetime(2026, 8, 1), datetime(2027, 8, 1)], 
                     'EDP SP':        [datetime(2026, 11, 1), datetime(2027, 11, 1)], 'Elektro':          [datetime(2026, 9, 1), datetime(2027, 9, 1)], 'Energisa':       [datetime(2026, 5, 1), datetime(2027, 5, 1)], 
                     'Equatorial':    [datetime(2026, 10, 1), datetime(2027, 10, 1)]}

ponto_de_corte = {'Celesc': 1, 'Cemig': 1, 'Copel': 4, 'CPFL': 2, 'CPFL Piratininga': 2, 'EDP ES': 1,'EDP SP': 1, 'Elektro': 2, 'Energisa': 3, 'Equatorial': 1}

# Demanda Ativa - Tarifa R$/kW / Demanda Isenta - Tarifa R$/kW / Demanda Geracao - Tarifa R$/kW / Consumo - R$ / Reativo R$/kW / Atraso Porcentagem % / Cosip R$
# Tarifas Novas = Tarifas Antiga * 6,05%

impostos = {
            'Celesc':           [0.0096, 0.0441, 0.17], 
            'Cemig':            [0.0125, 0.0575, 0.18], 
            'Copel':            [0.0122, 0.0563, 0.19], 
            'CPFL':             [0.0122, 0.0563, 0.18], 
            'CPFL Piratininga': [0.0097, 0.0451, 0.18],
            'EDP ES':           [0.0066, 0.0302, 0.17],
            'EDP SP':           [0.0097, 0.0451, 0.18], 
            'Elektro':          [0.0091, 0.0417, 0.18], 
            'Energisa':         [0.0165, 0.0760, 0.17], 
            'Equatorial':       [0.0039, 0.0182, 0.19]
            }

tarifa = {  
            'Celesc':           {          
                                 'A':  {'Demanda Ativa': 18.09,      'Demanda Ativa Isenta': 18.09,     'Demanda Geracao': 4.14,     'Demanda Geracao Isenta': 4.14,       'Consumo': 90, 'Reativo': 0.32193,   'Atraso': 2, 'Cosip': 50},
                                 'N':  {'Demanda Ativa': 19.184445,  'Demanda Ativa Isenta': 19.184445, 'Demanda Geracao': 4.390470, 'Demanda Geracao Isenta': 4.390470,   'Consumo': 90, 'Reativo': 0.3414067, 'Atraso': 2, 'Cosip': 50},
                                 'NN': {'Demanda Ativa': 20.345104,  'Demanda Ativa Isenta': 20.345104, 'Demanda Geracao': 4.656093, 'Demanda Geracao Isenta': 4.656093,   'Consumo': 90, 'Reativo': 0.3620618, 'Atraso': 2, 'Cosip': 50}
                                 },

            'Cemig':            {
                                 'A':  {'Demanda Ativa': 22.81 ,     'Demanda Ativa Isenta': 22.81,     'Demanda Geracao': 13.01,     'Demanda Geracao Isenta': 13.01,      'Consumo': 0, 'Reativo': 0.31728,  'Atraso': 2, 'Cosip': 0},
                                 'N':  {'Demanda Ativa': 24.190005,  'Demanda Ativa Isenta': 24.190005, 'Demanda Geracao': 13.797105, 'Demanda Geracao Isenta': 13.797105,  'Consumo': 0, 'Reativo': 0.336475, 'Atraso': 2, 'Cosip': 0},
                                 'NN': {'Demanda Ativa': 25.653495,  'Demanda Ativa Isenta': 25.653495, 'Demanda Geracao': 14.631829, 'Demanda Geracao Isenta': 14.631829,  'Consumo': 0, 'Reativo': 0.356832, 'Atraso': 2, 'Cosip': 0}
                                 },

            'Copel':            {
                                 'A':  {'Demanda Ativa': 20.78,      'Demanda Ativa Isenta': 20.78,    'Demanda Geracao': 7.75,     'Demanda Geracao Isenta': 7.75,       'Consumo': 90, 'Reativo': 0.275750, 'Atraso': 2, 'Cosip': 50},
                                 'N':  {'Demanda Ativa': 22.037190,  'Demanda Ativa Isenta': 22.03719, 'Demanda Geracao': 8.218875, 'Demanda Geracao Isenta': 8.218875,   'Consumo': 90, 'Reativo': 0.292433, 'Atraso': 2, 'Cosip': 50},
                                 'NN': {'Demanda Ativa': 23.370439,  'Demanda Ativa Isenta': 22.03719, 'Demanda Geracao': 8.716117, 'Demanda Geracao Isenta': 8.716117,   'Consumo': 90, 'Reativo': 0.310125, 'Atraso': 2, 'Cosip': 50}
                                 },

            'CPFL':             {
                                 'A':  {'Demanda Ativa': 15.93,     'Demanda Ativa Isenta': 15.93,     'Demanda Geracao': 7.08,      'Demanda Geracao Isenta': 7.08,      'Consumo': 90, 'Reativo': 0.2902,   'Atraso': 2, 'Cosip': 0},
                                 'N':  {'Demanda Ativa': 16.893765, 'Demanda Ativa Isenta': 16.893765, 'Demanda Geracao': 7.508340,  'Demanda Geracao Isenta': 7.508340,  'Consumo': 90, 'Reativo': 0.307757, 'Atraso': 2, 'Cosip': 0},
                                 'NN': {'Demanda Ativa': 17.915837, 'Demanda Ativa Isenta': 17.915837, 'Demanda Geracao': 7.962594,  'Demanda Geracao Isenta': 7.962594,  'Consumo': 90, 'Reativo': 0.326376, 'Atraso': 2, 'Cosip': 0}
                                 },

            'CPFL Piratininga': {
                                 'A':  {'Demanda Ativa': 15.93,     'Demanda Ativa Isenta': 15.93,     'Demanda Geracao': 7.08,     'Demanda Geracao Isenta': 7.08,      'Consumo': 90, 'Reativo': 0.2902,   'Atraso': 2, 'Cosip': 0},
                                 'N':  {'Demanda Ativa': 16.893765, 'Demanda Ativa Isenta': 16.893765, 'Demanda Geracao': 7.508340, 'Demanda Geracao Isenta': 7.508340,  'Consumo': 90, 'Reativo': 0.307757, 'Atraso': 2, 'Cosip': 0},
                                 'NN': {'Demanda Ativa': 17.915837, 'Demanda Ativa Isenta': 17.915837, 'Demanda Geracao': 7.962594, 'Demanda Geracao Isenta': 7.962594,  'Consumo': 90, 'Reativo': 0.326376, 'Atraso': 2, 'Cosip': 0}
                                 },

            'EDP ES':           {
                                 'A':  {'Demanda Ativa': 31.93,     'Demanda Ativa Isenta': 31.93,     'Demanda Geracao': 12.31,     'Demanda Geracao Isenta': 12.31,    'Consumo': 90, 'Reativo': 0.3206800, 'Atraso': 2, 'Cosip': 50},
                                 'N':  {'Demanda Ativa': 33.87,     'Demanda Ativa Isenta': 33.87,     'Demanda Geracao': 12.61,     'Demanda Geracao Isenta': 12.61,    'Consumo': 90, 'Reativo': 0.3499968, 'Atraso': 2, 'Cosip': 50},
                                 'NN': {'Demanda Ativa': 34.08,     'Demanda Ativa Isenta': 34.08,     'Demanda Geracao': 13.37,     'Demanda Geracao Isenta': 13.37,    'Consumo': 90, 'Reativo': 0.3711716, 'Atraso': 2, 'Cosip': 50}
                                },

            'EDP SP':           {
                                 'A':  {'Demanda Ativa': 12.02,     'Demanda Ativa Isenta': 12.02,     'Demanda Geracao': 4.11,     'Demanda Geracao Isenta': 4.11,      'Consumo': 90, 'Reativo': 0.33003,   'Atraso': 2, 'Cosip': 50},
                                 'N':  {'Demanda Ativa': 12.747210, 'Demanda Ativa Isenta': 12.747210, 'Demanda Geracao': 4.253665, 'Demanda Geracao Isenta': 4.253665,  'Consumo': 90, 'Reativo': 0.3499968, 'Atraso': 2, 'Cosip': 50},
                                 'NN': {'Demanda Ativa': 13.518416, 'Demanda Ativa Isenta': 13.518416, 'Demanda Geracao': 4.511001, 'Demanda Geracao Isenta': 4.511001,  'Consumo': 90, 'Reativo': 0.3711716, 'Atraso': 2, 'Cosip': 50}
                                 },

            'Elektro':          {
                                 'A':  {'Demanda Ativa': 28.61,     'Demanda Ativa Isenta': 28.61,     'Demanda Geracao': 14.06,     'Demanda Geracao Isenta': 14.06,     'Consumo': 90, 'Reativo': 0.3182, 'Atraso': 2, 'Cosip': 0},
                                 'N':  {'Demanda Ativa': 25.20,     'Demanda Ativa Isenta': 25.20,     'Demanda Geracao': 11.59,     'Demanda Geracao Isenta': 11.59,     'Consumo': 90, 'Reativo': 0.3263, 'Atraso': 2, 'Cosip': 0},
                                 'NN': {'Demanda Ativa': 26.72,     'Demanda Ativa Isenta': 26.72,     'Demanda Geracao': 12.29,     'Demanda Geracao Isenta': 12.29,     'Consumo': 90, 'Reativo': 0.3460, 'Atraso': 2, 'Cosip': 0}
                                 },

            'Energisa':         {
                                 'A':  {'Demanda Ativa': 32.74,    'Demanda Ativa Isenta': 32.74,    'Demanda Geracao': 17.24, 'Demanda Geracao Isenta': 17.24, 'Consumo': 0, 'Reativo': 0.302660, 'Atraso': 2, 'Cosip': 0},
                                 'N':  {'Demanda Ativa': 34.03,    'Demanda Ativa Isenta': 34.03,    'Demanda Geracao': 20.65, 'Demanda Geracao Isenta': 20.65, 'Consumo': 0, 'Reativo': 0.294140, 'Atraso': 2, 'Cosip': 0},
                                 'NN': {'Demanda Ativa': 36.82137, 'Demanda Ativa Isenta': 36.82137, 'Demanda Geracao': 23.90, 'Demanda Geracao Isenta': 23.90, 'Consumo': 0, 'Reativo': 0.294140, 'Atraso': 2, 'Cosip': 0}
                                 },

            'Equatorial':       {
                                 'A':  {'Demanda Ativa': 26.08,     'Demanda Ativa Isenta': 26.08,     'Demanda Geracao': 13.01,     'Demanda Geracao Isenta': 13.01,      'Consumo': 0, 'Reativo': 0.32414,  'Atraso': 2, 'Cosip': 0},
                                 'N':  {'Demanda Ativa': 27.657840, 'Demanda Ativa Isenta': 27.657840, 'Demanda Geracao': 13.797105, 'Demanda Geracao Isenta': 13.797105,  'Consumo': 0, 'Reativo': 0.343750, 'Atraso': 2, 'Cosip': 0},
                                 'NN': {'Demanda Ativa': 29.331139, 'Demanda Ativa Isenta': 29.331139, 'Demanda Geracao': 14.631829, 'Demanda Geracao Isenta': 14.631829,  'Consumo': 0, 'Reativo': 0.364547, 'Atraso': 2, 'Cosip': 0}
                                 }
                                 }

estrutura_UGs = {0.5:   [0.5],
                 0.85:   [1],
                 1:     [1],
                 1.125: [1, 0.125],
                 1.5:   [1, 0.5],
                 2:     [1, 1],
                 2.25:  [1, 1, 0.25],
                 2.5:   [1, 1, 0.5],
                 3:     [1, 1, 1],
                 4:     [1, 1, 1, 1],
                 5:     [1, 1, 1, 1, 1]}

# 0.125 - 0.250 - 0.500 - 1.000
# 'Copel':            {0.125: [6, 119, 0, 0, 90, 268, 0.02, 50], 0.25: [6, 244, 0, 0, 90, 268, 0.02, 50],  0.5: [6, 494, 0, 0, 90, 268, 0.02, 50],  1: [6, 994, 0, 0, 90, 268, 0.02, 50]}
# 'Copel':            {0.125: [6, 24, 95, 0, 90, 268, 0.02, 50], 0.25: [6, 24, 220, 0, 90, 268, 0.02, 50], 0.5: [6, 24, 470, 0, 90, 268, 0.02, 50], 1: [6, 24, 970, 0, 90, 268, 0.02, 50]}
medicao_e_contrato = {
                      'Celesc':           {0.125: [5, 25, 95, 0, 90, 500, 0.02, 50], 0.25: [5, 25, 220, 0, 90, 500, 0.02, 50], 0.5: [5, 25, 470, 0, 90, 500, 0.02, 50], 1: [5, 25, 970, 0, 90, 500, 0.02, 50]}, 

                      'Cemig':            {0.125: [5, 25, 95, 0, 0, 765, 0.02, 0],   0.25: [5, 25, 220, 0, 0, 765, 0.02, 0],   0.5: [5, 25, 470, 0, 0, 765, 0.02, 0],   1: [5, 25, 970, 0, 0, 765, 0.02, 0]},

                      'Copel':            {0.125: [6, 24, 95, 0, 90, 268, 0.02, 50], 0.25: [6, 24, 220, 0, 90, 268, 0.02, 50], 0.5: [6, 24, 470, 0, 90, 268, 0.02, 50], 1: [6, 24, 970, 0, 90, 268, 0.02, 50]}, 

                      'CPFL':             {0.125: [5, 25, 95, 0, 90, 600, 0.02, 0],  0.25: [5, 25, 220, 0, 90, 600, 0.02, 0],  0.5: [5, 25, 470, 0, 90, 600, 0.02, 0],  1: [5, 25, 970, 0, 90, 600, 0.02, 0]}, 

                      'CPFL Piratininga': {0.125: [5, 25, 95, 0, 90, 600, 0.02, 0],  0.25: [5, 25, 220, 0, 90, 600, 0.02, 0],  0.5: [5, 25, 470, 0, 90, 600, 0.02, 0],  1: [5, 25, 970, 0, 90, 600, 0.02, 0]},

                      'EDP ES':           {0.125: [4, 26, 95, 0, 90, 350, 0.02, 0],  0.25: [4, 26, 220, 0, 90, 350, 0.02, 0],  0.5: [4, 26, 470, 0, 90, 350, 0.02, 0],  1: [4, 26, 970, 0, 90, 350, 0.02, 0]}, 

                      'EDP SP':           {0.125: [5, 25, 95, 0, 90, 400, 0.02, 0],  0.25: [5, 25, 220, 0, 90, 400, 0.02, 0],  0.5: [5, 25, 470, 0, 90, 400, 0.02, 0],  1: [5, 25, 970, 0, 90, 400, 0.02, 0]}, 

                      'Elektro':          {0.125: [4, 26, 95, 0, 90, 588, 0.02, 0],  0.25: [4, 26, 220, 0, 90, 588, 0.02, 0],  0.5: [4, 26, 470, 0, 90, 588, 0.02, 0],  1: [4, 26, 970, 0, 90, 588, 0.02, 0]}, 

                      'Energisa':         {0.125: [7, 23, 95, 0, 0, 239, 0.02, 0],   0.25: [7, 23, 220, 0, 0, 239, 0.02, 0],   0.5: [7, 23, 470, 0, 0, 239, 0.02, 0],   1: [7, 23, 970, 0, 0, 239, 0.02, 0]}, 

                      'Equatorial':       {0.125: [5, 25, 95, 0, 0, 650, 0.02, 0],   0.25: [5, 25, 220, 0, 0, 650, 0.02, 0],   0.5: [5, 25, 470, 0, 0, 650, 0.02, 0],   1: [5, 25, 970, 0, 0, 650, 0.02, 0]}
                      }

# Caminhos
caminho_arquivo = r'C:\Users\lucas.conceicao\Relatorio de Compensacao\Diversos\Conexao_Usinas.xlsx'
caminho_analise = 'C:/Users/lucas.conceicao/Relatorio de Compensacao/Diversos/CUSD_UFVs.xlsx'

# Dataframes
dataframe_plan_conex = None
dataframe_plan_conex_ajus = None
dataframe_preco_cmpt = None
dataframe_previsao_de_custo = None

def leitor():
    
    global dataframe_plan_conex

    dataframe_plan_conex = pandas.read_excel(caminho_arquivo, sheet_name='Planejamento de Conexao')

def tratamento_dataframe_plan_conex():
    
    global dataframe_plan_conex_ajus

    # Tratamento colunas - Exclusão
    dataframe_plan_conex_ajus = dataframe_plan_conex.drop(columns=['MWp', 'P50', 'P90', 'Conexão'])
    
    # Tratamento colunas - Ordenação
    dataframe_plan_conex_ajus = dataframe_plan_conex_ajus[['Cluster', 'Distribuidora', 'Usina', 'MWac', 'Data Ajustada']]

    dataframe_plan_conex_ajus.loc[dataframe_plan_conex_ajus['MWac'] == 1.98, 'MWac'] = 2

def dataframe_preco_competencia(data_inicial, data_final):
    
    global dataframe_preco_cmpt

    dataframe_preco_cmpt = pandas.DataFrame(columns=['Data', 'Distribuidora', 'Demanda Ativa', 'Demanda Ativa Isenta', 'Demanda Geracao', 'Demanda Geracao Isenta', 'Reativo'])

    for distribuidora in distribuidoras:

        data_corrente = data_inicial

        while data_corrente <= data_final:

            if data_corrente > revisao_tarifaria[distribuidora][1]:
                
                data = data_corrente
                concessionaria = distribuidora
                demanda_ativa = tarifa[distribuidora]['NN']['Demanda Ativa'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1]))) * (1 / (1 - impostos[distribuidora][2]))
                demanda_Ativa_isenta = tarifa[distribuidora]['NN']['Demanda Ativa Isenta'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1])))
                if distribuidora in ['Celesc', 'Cemig', 'Copel', 'EDP ES', 'Elektro', 'Equatorial']:
                    demanda_geracao = tarifa[distribuidora]['NN']['Demanda Geracao'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1]))) * (1 / (1 - impostos[distribuidora][2]))
                else:
                    demanda_geracao = tarifa[distribuidora]['NN']['Demanda Geracao'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1])))
                demanda_geracao_isenta = tarifa[distribuidora]['NN']['Demanda Geracao Isenta'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1])))
                reativo = tarifa[distribuidora]['NN']['Reativo'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1]))) * (1 / (1 - impostos[distribuidora][2])) 

            elif data_corrente > revisao_tarifaria[distribuidora][0]:

                data = data_corrente
                concessionaria = distribuidora
                demanda_ativa = tarifa[distribuidora]['N']['Demanda Ativa'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1]))) * (1 / (1 - impostos[distribuidora][2]))
                demanda_Ativa_isenta = tarifa[distribuidora]['N']['Demanda Ativa Isenta'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1])))
                if distribuidora in ['Celesc', 'Cemig', 'Copel', 'EDP ES', 'Elektro', 'Equatorial']:
                    demanda_geracao = tarifa[distribuidora]['N']['Demanda Geracao'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1]))) * (1 / (1 - impostos[distribuidora][2]))
                else:
                    demanda_geracao = tarifa[distribuidora]['N']['Demanda Geracao'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1])))
                demanda_geracao_isenta = tarifa[distribuidora]['N']['Demanda Geracao Isenta'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1])))
                reativo = tarifa[distribuidora]['N']['Reativo'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1]))) * (1 / (1 - impostos[distribuidora][2])) 

            else:

                data = data_corrente
                concessionaria = distribuidora
                demanda_ativa = tarifa[distribuidora]['A']['Demanda Ativa'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1]))) * (1 / (1 - impostos[distribuidora][2]))
                demanda_Ativa_isenta = tarifa[distribuidora]['A']['Demanda Ativa Isenta'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1])))
                if distribuidora in ['Celesc', 'Cemig', 'Copel', 'EDP ES', 'Elektro', 'Equatorial']:
                    demanda_geracao = tarifa[distribuidora]['A']['Demanda Geracao'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1]))) * (1 / (1 - impostos[distribuidora][2]))
                else:
                    demanda_geracao = tarifa[distribuidora]['A']['Demanda Geracao'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1])))
                demanda_geracao_isenta = tarifa[distribuidora]['A']['Demanda Geracao Isenta'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1])))
                reativo = tarifa[distribuidora]['A']['Reativo'] * (1 / (1-(impostos[distribuidora][0] + impostos[distribuidora][1]))) * (1 / (1 - impostos[distribuidora][2]))

            dataframe_preco_cmpt.loc[len(dataframe_preco_cmpt)] = [data, concessionaria, demanda_ativa, demanda_Ativa_isenta, demanda_geracao, demanda_geracao_isenta, reativo]

            data_corrente = data_corrente + relativedelta(months=1)

    #print(dataframe_preco_cmpt)

def previsao_de_custo(data_inicial, data_final):

    global dataframe_previsao_de_custo

    dataframe_previsao_de_custo = pandas.DataFrame(columns=['Data', 'Cluster', 'Distribuidora', 'Usina', 'MWac', 'Demanda Ativa', 'Demanda Ativa Isenta', 'Demanda Geração', 'Demanda Geração Isenta', 'Consumo', 'Reativo', 'Atraso', 'Cosip'])

    for index, row in dataframe_plan_conex_ajus.iterrows():

        data_corrente = data_inicial
        contador_mes = int((data_inicial.year - row['Data Ajustada'].year) * 12 + (data_inicial.month - row['Data Ajustada'].month))
        mes_de_corte = ponto_de_corte[row['Distribuidora']]
        
        while data_corrente <= data_final:
            
            if row['Data Ajustada'] >= data_corrente:
                pass
            else:

                base_precos = dataframe_preco_cmpt[(dataframe_preco_cmpt['Data']==data_corrente) & (dataframe_preco_cmpt['Distribuidora']==row['Distribuidora'])]
                preco_d_a = base_precos['Demanda Ativa'].values[0]
                preco_d_a_i = base_precos['Demanda Ativa Isenta'].values[0]
                preco_d_g = base_precos['Demanda Geracao'].values[0]
                preco_d_g_i = base_precos['Demanda Geracao Isenta'].values[0]
                preco_r = base_precos['Reativo'].values[0]

                if contador_mes < mes_de_corte:
                    
                    for ug in estrutura_UGs[row['MWac']]:
 
                        data = data_corrente
                        cluster = row['Cluster']
                        distribuidora = row['Distribuidora']
                        usina = row['Usina']
                        potencia = ug
                        demanda_ativa = 5 * preco_d_a
                        demanda_ativa_isenta = 25 * preco_d_a_i
                        demanda_geracao = 0 * preco_d_g
                        demanda_geracao_isenta = 0 * preco_d_g_i
                        consumo = 90
                        reativo = 0 * preco_r
                        atraso = 0
                        cosip = 50

                        dataframe_previsao_de_custo.loc[len(dataframe_previsao_de_custo)] = [data, cluster, distribuidora, usina, potencia, demanda_ativa, demanda_ativa_isenta, demanda_geracao, demanda_geracao_isenta, consumo, reativo, atraso, cosip]

                else:

                    for ug in estrutura_UGs[row['MWac']]:

                        medidas = medicao_e_contrato[row['Distribuidora']][ug]

                        data = data_corrente
                        cluster = row['Cluster']
                        distribuidora = row['Distribuidora']
                        usina = row['Usina']
                        potencia = ug
                        demanda_ativa = round(medidas[0] * preco_d_a, 0)
                        demanda_ativa_isenta = round(medidas[1] * preco_d_a_i, 0)
                        demanda_geracao = round(medidas[2] * preco_d_g, 0)
                        demanda_geracao_isenta = round(medidas[3] * preco_d_g_i, 0)
                        consumo = round(medidas[4], 0)
                        reativo = round(medidas[5] * preco_r, 0)
                        total = round(demanda_ativa + demanda_ativa_isenta + demanda_geracao + demanda_geracao_isenta, 0)
                        atraso = round(medidas[6] * total, 0)
                        cosip = round(medidas[7], 0)

                        dataframe_previsao_de_custo.loc[len(dataframe_previsao_de_custo)] = [data, cluster, distribuidora, usina, potencia, demanda_ativa, demanda_ativa_isenta, demanda_geracao, demanda_geracao_isenta, consumo, reativo, atraso, cosip]

            contador_mes += 1
            data_corrente = data_corrente + relativedelta(months=1)

    dataframe_previsao_de_custo = dataframe_previsao_de_custo.sort_values(by=['Data'], ascending=[True])
    dataframe_previsao_de_custo['Data'] = dataframe_previsao_de_custo['Data'].dt.strftime('%m/%y')

def salvar():
    
    dataframe_previsao_de_custo.to_excel(r"C:\Users\lucas.conceicao\Relatorio de Compensacao\Previsão de Custo com Faturas.xlsx")

def principal(data_inicial, data_final):
    
    leitor()

    tratamento_dataframe_plan_conex()
    
    dataframe_preco_competencia(data_inicial, data_final)

    previsao_de_custo(data_inicial, data_final)
    
    print(dataframe_previsao_de_custo)

    salvar()

principal(datetime(2026, 6, 1), datetime(2027, 12, 1))