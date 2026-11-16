import pandas
import os
from datetime import datetime

# Caminhos
caminho_arquivo = 'C:/Users/lucas.conceicao/Downloads/data.xlsx'
caminho_analise = 'C:/Users/lucas.conceicao/Downloads'

# Dataframes
dataframe = None
dataframe_pronto = None
dataframe_faturas_vigentes = None

def leitor():
    
    global dataframe

    dataframe = pandas.read_excel(caminho_arquivo, engine='openpyxl')

    return dataframe

def tratamento_dataframe():
    
    global dataframe_pronto
    # Tratamento colunas - Exclusão
    dataframe_pronto = dataframe.drop(columns=['Recurso', 'MWp', 'MWac', 'P50 Anual', 'P90 Anual', 'Tipo de Projeto', 'Observação'])

    # Tratamento colunas - Ajuste de nomes
    dataframe_pronto.rename(columns={'Recurso.1': 'Recurso'}, inplace=True)

    # Tratamento colunas - Ajuste de posição
    dataframe_pronto = dataframe_pronto[['Recurso', 'UF', 'Concessionária', 'UFV', 'Conexão Prevista']]

    # Tratamento linhas - Exclusão
    dataframe_pronto.drop([0] + list(dataframe_pronto.tail(3).index), inplace=True)
    
    dataframe_pronto = dataframe_pronto[dataframe_pronto['Conexão Prevista'] < datetime.today()]
   
    return dataframe_pronto

def faturas_vigentes():

    global dataframe_faturas_vigentes

    dataframe_faturas_vigentes = dataframe_pronto

    dataframe_faturas_vigentes['Quantidade de Faturas'] = (datetime.today().year - dataframe_faturas_vigentes['Conexão Prevista'].dt.year) * 12 + (datetime.today().month - dataframe_faturas_vigentes['Conexão Prevista'].dt.month)

    dataframe_faturas_vigentes['Conexão Prevista'] = dataframe_faturas_vigentes['Conexão Prevista'].dt.strftime('%d/%b/%y')

    dataframe_faturas_vigentes = dataframe_faturas_vigentes.astype({'Recurso': 'int32'})
    
    dataframe_faturas_vigentes.sort_values('Conexão Prevista')

    return dataframe_faturas_vigentes

def salvar():

    caminho_arquivo_final = os.path.join(caminho_analise, 'Faturas_Vigentes.xlsx')

    if not os.path.exists(caminho_arquivo_final):
        dataframe_faturas_vigentes.to_excel(caminho_arquivo_final, index=False)

def principal():
    
    leitor()
    
    tratamento_dataframe()

    faturas_vigentes()

    salvar()

principal()