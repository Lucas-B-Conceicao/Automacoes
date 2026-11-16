from tabula import read_pdf
import PyPDF2
import pandas
import regex
import os
import shutil
from openpyxl import load_workbook
from tabulate import tabulate
import calendar
from datetime import datetime

caminho_pdf = 'C:/Users/lucas.conceicao/Downloads/DemonstrativoMicroMiniGeracao.pdf'
caminho_depara = 'C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/03.Análises/DEPARA_COPEL.xlsx'
caminho_analise_de_faturas = 'C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Faturas/Análise de Faturas.xlsm'
caminho_lista_vigente = 'C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/Arquivos BI/Controle de Estoque - USINAS/Base previsão faturamento - Lista Vigente.xlsm'
caminho_recusados = "C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento/Extras/Recusados.xlsx"
caminho_compensados = "C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento/Extras/Sem_compensacao.xlsx"
distribuidoras = {'ELEKTRO': ['45709033', '45709254', '45709360', '45963592', '45963894', '45948194', '45948160', '45948321', '45878765', '45869456', '45869324', '45948046'], }
dicio_meses = {'JAN': '1', 'FEV': '2', 'MAR': '3', 'ABR': '4', 'MAI': '5', 'JUN': '6', 'JUL': '7', 'AGO': '8', 'SET': '9', 'OUT': '10', 'NOV': '11', 'DEZ': '12'}
dicio_ano = {'24': '2024', '25': '2025', '26': '2026'}

arquivo_aberto = None

recusadas = []
energia_recusadas = []
energia_recusada = None
distribuidora = None
competencia = None
ug = None
ug_nova = None

geracao_ug = None
consumo_ug = None
transferencia_ug = None

percentual_recusados = None

geracao_calc = None
consumo_calc = None
compensado_calc = None
transferido_calc = None

dataframe_deapara = pandas.read_excel(caminho_analise_de_faturas, sheet_name="NOVOS No. UG")
dataframe_deapara = dataframe_deapara[["Antiga UG", "Nova UG"]]
de_para_dict = dict(zip(dataframe_deapara["Antiga UG"],dataframe_deapara['Nova UG']))
dataframe_analise = None
dataframe_tarifas = None
dataframe_aprovados = None
dataframe_recusados = None
sem_compensacao = None
dataframe_divergencias = None

def leitor_arquivo():

  global caminho_pdf
  global arquivo_aberto

  arquivo_aberto = open(caminho_pdf, 'rb')
  arquivo_lido = PyPDF2.PdfReader(arquivo_aberto)
  quantidade_paginas = len(arquivo_lido.pages)

  return arquivo_lido, quantidade_paginas

def unidade_geradora(pagina, arquivo):

  global caminho_pdf
  global distribuidoras
  global distribuidora
  global competencia
  global ug
  global ug_nova

  pagina_lida = arquivo.pages[pagina]
  conteudo = pagina_lida.extract_text()
  ug = encontrar_uc = regex.findall(r'UC Conces.\.?\s*:\s*(\d+)', conteudo)[0]
  competencia = regex.findall(r'Emitido em\s*:\s(\d{2}/\d{2}/\d{4})', conteudo)[0][3:].replace("/", "-")

  for D in distribuidoras.items():
    if encontrar_uc in D[1]:
      distribuidora = D[0]

  tabelas = read_pdf(caminho_pdf, pages=pagina+1, multiple_tables=True, lattice=True, pandas_options={'header': 0})
  tabelas = tabelas[1]
  tabelas.to_csv('{}.csv'.format(encontrar_uc), encoding='utf-8')

  return encontrar_uc

def check_empty(pagina):

  global caminho_pdf

  tabelas = read_pdf(caminho_pdf, pages=pagina+1, multiple_tables=True, lattice=True, pandas_options={'header': 0})
  tabelas = tabelas[1]

  return tabelas.empty

def check_columns(pagina):

  global caminho_pdf

  tabelas = read_pdf(caminho_pdf, pages=pagina+1, multiple_tables=True, lattice=True, pandas_options={'header': 0})
  tabelas = tabelas[1]
  
  return tabelas.shape[1]

def unidade_consumidora(pagina, arquivo):

  global caminho_pdf

  pagina_lida = arquivo.pages[pagina]
  conteudo = pagina_lida.extract_text()
  encontrar_uc = regex.findall(r'UC Conces.\.?\s*:\s*(\d+)', conteudo)[0]
  tabelas = read_pdf(caminho_pdf, pages=pagina+1, multiple_tables=True, lattice=True, pandas_options={'header': 0})
  tabelas = tabelas[1]
  tabelas.to_csv('{}.csv'.format(encontrar_uc), encoding='utf-8')

  return encontrar_uc

def dataframe_pandas_copel(uc, pagina):
  
  df = pandas.read_csv('{}.csv'.format(uc), sep=',')

  if df.shape[1] == 25:
    df.rename(columns={'Unnamed: 0.1': 'Indice',
                    'Referência': 'Referência',
                    
                    'Saldo Anterior': 'Saldo Anterior PT', 
                    'Cred. Receb. Outra UC': 'Saldo Anterior FP', 
                    'Energia Injetada': 'Saldo Anterior TP', 

                    'Energia Ativa': 'Cred. Receb. Outra UC PT', 
                    'Credito Utilizado no Mês': 'Cred. Receb. Outra UC FP', 
                    'Percentual': 'Cred. Receb. Outra UC TP', 

                    'Prioridade': 'Energia Injetada PT', 
                    'Saldo Transferido Outra UC': 'Energia Injetada FP',
                    'Saldo Final': 'Energia Injetada TP',

                    'Unnamed: 0': 'Energia Ativa PT', 
                    'Unnamed: 1': 'Energia Ativa FP', 
                    'Unnamed: 2': 'Energia Ativa TP', 

                    'Unnamed: 3': 'Credito Utilizado no Mês PT', 
                    'Unnamed: 4': 'Credito Utilizado no Mês FP', 
                    'Unnamed: 5': 'Credito Utilizado no Mês TP',

                    'Unnamed: 6': 'Percentual', 
                    'Unnamed: 7': 'Prioridade', 

                    'Unnamed: 8': 'Saldo Transferido Outra UC PT', 
                    'Unnamed: 9': 'Saldo Transferido Outra UC FP', 
                    'Unnamed: 10': 'Saldo Transferido Outra UC TP',

                    'Unnamed: 11': 'Saldo Final PT', 
                    'Unnamed: 12': 'Saldo Final FP', 
                    'Unnamed: 13': 'Saldo Final TP'},
                        inplace=True)
    
    dados_coluna_uc = [uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc]
    df.insert(1, 'Código', dados_coluna_uc)

    df.drop(df.head(1).index, inplace=True)
    df.drop(df.columns[0], axis=1, inplace=True)

    os.remove('{}.csv'.format(uc))   

  else:
      if pagina == 0:
        df.rename(columns={'Unnamed: 0.1': 'Indice',
                            'Referência': 'Referência',
                            'Saldo Anterior': 'Saldo Anterior PT', 
                            'Cred. Receb. Outra UC': 'Saldo Anterior FP',
                            'Energia Injetada': 'Cred. Receb. Outra UC PT', 
                            'Energia Ativa': 'Cred. Receb. Outra UC FP',
                            'Credito Utilizado no Mês': 'Energia Injetada PT', 
                            'Percentual': 'Energia Injetada FP',
                            'Prioridade': 'Energia Ativa PT', 
                            'Saldo Transferido Outra UC': 'Energia Ativa FP',
                            'Saldo Final': 'Credito Utilizado no Mês PT',
                            'Unnamed: 0': 'Credito Utilizado no Mês FP',
                            'Unnamed: 1': 'Percentual',
                            'Unnamed: 2': 'Prioridade', 
                            'Unnamed: 3': 'Saldo Transferido Outra UC PT', 
                            'Unnamed: 4': 'Saldo Transferido Outra UC FP',
                            'Unnamed: 5': 'Saldo Final PT', 
                            'Unnamed: 6': 'Saldo Final FP'},
                        inplace=True)
      
        dados_nova_coluna = [0,0,0,0,0,0,0,0,0,0,0,0,0,0]
        df.insert(4, 'Saldo Anterior TP', dados_nova_coluna)
        df.insert(7, 'Cred. Receb. Outra UC TP', dados_nova_coluna)
        df.insert(10, 'Energia Injetada TP', dados_nova_coluna)
        df.insert(13, 'Energia Ativa TP', dados_nova_coluna)
        df.insert(16, 'Credito Utilizado no Mês TP', dados_nova_coluna)
        df.insert(21, 'Saldo Transferido Outra UC TP', dados_nova_coluna)
        df.insert(24, 'Saldo Final TP', dados_nova_coluna)
      else:
        df.rename(columns={'Unnamed: 0.1': 'Indice',
                            'Referência': 'Referência',
                            'Saldo Anterior': 'Saldo Anterior FP', 
                            'Cred. Receb. Outra UC': 'Saldo Anterior TP',
                            'Energia Injetada': 'Cred. Receb. Outra UC FP', 
                            'Energia Ativa': 'Cred. Receb. Outra UC TP',
                            'Credito Utilizado no Mês': 'Energia Injetada FP', 
                            'Percentual': 'Energia Injetada TP',
                            'Prioridade': 'Energia Ativa FP', 
                            'Saldo Transferido Outra UC': 'Energia Ativa TP',
                            'Saldo Final': 'Credito Utilizado no Mês FP',
                            'Unnamed: 0': 'Credito Utilizado no Mês TP',
                            'Unnamed: 1': 'Percentual',
                            'Unnamed: 2': 'Prioridade', 
                            'Unnamed: 3': 'Saldo Transferido Outra UC FP', 
                            'Unnamed: 4': 'Saldo Transferido Outra UC TP',
                            'Unnamed: 5': 'Saldo Final FP', 
                            'Unnamed: 6': 'Saldo Final TP'},
                        inplace=True)
        
        dados_nova_coluna = [0,0,0,0,0,0,0,0,0,0,0,0,0,0]
        df.insert(2, 'Saldo Anterior PT', dados_nova_coluna)
        df.insert(5, 'Cred. Receb. Outra UC PT', dados_nova_coluna)
        df.insert(8, 'Energia Injetada PT', dados_nova_coluna)
        df.insert(11, 'Energia Ativa PT', dados_nova_coluna)
        df.insert(14, 'Credito Utilizado no Mês PT', dados_nova_coluna)
        df.insert(19, 'Saldo Transferido Outra UC PT', dados_nova_coluna)
        df.insert(22, 'Saldo Final PT', dados_nova_coluna)

      dados_coluna_uc = [uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc]
      df.insert(1, 'Código', dados_coluna_uc)

      df.drop(df.head(1).index, inplace=True)
      df.drop(df.columns[0], axis=1, inplace=True)

      os.remove('{}.csv'.format(uc))

  return df

def dataframe_analise_de_faturas(data):

  global dataframe_analise

  dataframe = pandas.read_excel(caminho_analise_de_faturas, sheet_name="Geral", skiprows=2)
  
  data_ajusta = f'{dicio_ano[data[3:5]]}-0{int(dicio_meses[data[0:3]])-1}-01'

  dataframe = dataframe[(dataframe["Mês"] == data_ajusta) & (dataframe["Distribuidora"] == "Elektro")]

  dataframe = dataframe[["Mês", "UG", "Cons P Reg", "Cons FP Reg", "Inj P Reg", "Inj FP Reg", "Saldo Anterior P kWh", "Saldo Anterior FP kWh",]]

  dataframe_analise = dataframe

  return dataframe_analise

def excedente_geracao(data):

  global dataframe_tarifas
  global geracao_calc
  global consumo_calc
  global compensado_calc
  global transferido_calc
  global energia_recusada
  global ug_nova

  ug_nova = de_para_dict[int(ug)]

  dataframe = dataframe_analise[dataframe_analise["UG"]==ug_nova]

  dataframe_tarifas = pandas.read_excel(caminho_analise_de_faturas, sheet_name="Tarifas")

  data_inicial = f'{dicio_ano[data[3:5]]}-0{int(dicio_meses[data[0:3]])-1}-01'
  data_final = f'{dicio_ano[data[3:5]]}-0{int(dicio_meses[data[0:3]])-1}-{calendar.monthrange(int(dicio_ano[data[3:5]]),int(dicio_meses[data[0:3]])-1)[1]}'

  te_ponta = dataframe_tarifas[(dataframe_tarifas["Início Vigência"] <= data_final) & (dataframe_tarifas["Fim Vigência"] >= data_inicial) & (dataframe_tarifas["Sigla"] == "Copel-DIS") 
                              & (dataframe_tarifas["Subgrupo"] == "A3a") & (dataframe_tarifas["Modalidade"] == "Verde") & (dataframe_tarifas["Detalhe"] == "Não se aplica") 
                              & (dataframe_tarifas["Posto"] != "Não se aplica") & (dataframe_tarifas["Unidade"] == "R$/MWh") & (dataframe_tarifas["Posto"] == "Ponta")]["TE"].values[0]

  te_fora_ponta = dataframe_tarifas[(dataframe_tarifas["Início Vigência"] <= data_final) & (dataframe_tarifas["Fim Vigência"] >= data_inicial) & (dataframe_tarifas["Sigla"] == "Copel-DIS") 
                                & (dataframe_tarifas["Subgrupo"] == "A3a") & (dataframe_tarifas["Modalidade"] == "Verde") & (dataframe_tarifas["Detalhe"] == "Não se aplica") 
                                & (dataframe_tarifas["Posto"] != "Não se aplica") & (dataframe_tarifas["Unidade"] == "R$/MWh") & (dataframe_tarifas["Posto"] == "Fora ponta")]["TE"].values[0]

  coeficiente_tarifario = te_ponta/te_fora_ponta

  consumo_ponta = dataframe["Cons P Reg"].values[0]
  consumo_fora_ponta = dataframe["Cons FP Reg"].values[0]

  injetado_ponta = dataframe["Inj P Reg"].values[0]
  injetado_fora_ponta = dataframe["Inj FP Reg"].values[0]

  banco_ponta = dataframe["Saldo Anterior P kWh"].values[0]
  banco_fora_ponta = dataframe["Saldo Anterior FP kWh"].values[0]

  consumo_calc = consumo_ponta + consumo_fora_ponta
  geracao_calc = injetado_ponta + injetado_fora_ponta
  

  if (injetado_ponta + banco_ponta) >= consumo_ponta and (injetado_fora_ponta + banco_fora_ponta) >= consumo_fora_ponta:

    credito_utilizado_ponta = consumo_ponta
    credito_utilizado_fora_ponta = consumo_fora_ponta

  elif (injetado_ponta + banco_ponta) < consumo_ponta and (injetado_fora_ponta + banco_fora_ponta) > consumo_fora_ponta:

    if (injetado_fora_ponta + banco_fora_ponta) - consumo_fora_ponta > ((consumo_ponta - injetado_ponta + banco_ponta) * coeficiente_tarifario):

      credito_utilizado_ponta = (injetado_ponta + banco_ponta) + ((consumo_ponta - injetado_ponta + banco_ponta) * coeficiente_tarifario)
      credito_utilizado_fora_ponta = consumo_fora_ponta

    else:

      credito_utilizado_ponta = (injetado_ponta + banco_ponta) + ((injetado_fora_ponta + banco_fora_ponta) - consumo_fora_ponta)
      credito_utilizado_fora_ponta = consumo_fora_ponta

  elif (injetado_ponta + banco_ponta) < consumo_ponta and (injetado_fora_ponta + banco_fora_ponta) <= consumo_fora_ponta:

      credito_utilizado_ponta = (injetado_ponta + banco_ponta)
      credito_utilizado_fora_ponta = (injetado_fora_ponta + banco_fora_ponta)

  compensado_calc = round(credito_utilizado_ponta + credito_utilizado_fora_ponta)

  excedente = round((injetado_ponta + injetado_fora_ponta) - compensado_calc)

  energia_recusada = round((percentual_recusados / 100) * excedente)

  transferido_calc = round(excedente - energia_recusada)

  return excedente

def dataframe_pandas_recebimento(data, exc):
  
  data_inicial = f'{dicio_ano[data[3:5]]}-0{int(dicio_meses[data[0:3]])-1}-01'
  
  dataframe_axs = pandas.read_excel(caminho_lista_vigente, "LISTA VIGÊNTE")

  dataframe_axs = dataframe_axs[(dataframe_axs["UG"] == int(ug)) & (dataframe_axs["COMPETÊNCIA"] == data_inicial)]

  dataframe_axs["Qtd_recebimento"] = dataframe_axs["Quota"] * exc

  return dataframe_axs

def dataframe_pandas_comparativo(df_copel, df_recebimento):
  
  global recusadas
  global energia_recusadas
  global dataframe_recusados

  dataframe_aprovados = pandas.DataFrame(columns=['Código', 'Quota', 'Recebimento AXS',
                                        'Recebimento Distribuidora', 'Diferenca', 'variacao'])

  for index, row in df_recebimento.iterrows():
      if row['UC'] in df_copel['Código'].values:
          codigo = row['UC']
          cota = row['Quota']
          Recebimento_AXS = int(row['Qtd_recebimento'])
          Recebimento_Distribuidora = int(df_copel[df_copel['Código'] == int(row['UC'])]['Cred. Receb. Outra UC FP'].values[0]) + int(df_copel[df_copel['Código'] == int(row['UC'])]['Cred. Receb. Outra UC PT'].values[0])
          Diferenca = Recebimento_AXS - Recebimento_Distribuidora
          variacao = Diferenca / Recebimento_Distribuidora * 100
          dataframe_aprovados.loc[len(dataframe_aprovados)] = [codigo, cota, Recebimento_AXS, Recebimento_Distribuidora, Diferenca, variacao]
      else:
          recusadas.append(row['UC'])
          energia_recusadas.append(row['Qtd_recebimento'])
  
  dataframe_recusados = pandas.DataFrame({'Comp':competencia, 'Distribuidora':'Elektro', 'Usina':'Usina', 'UG':ug, 'UC': recusadas, 'Motivo': 'Sem Informação', 'Energia': energia_recusadas})

  return dataframe_aprovados

def validacao_medicoes(dataframe):
   
  global geracao_ug
  global consumo_ug
  global carteira_anterior
  global carteira_atual
  global transferencia_ug

  geracao_ug = int(dataframe[dataframe['Código'] == int(ug)]['Energia Injetada FP'].values[0]) + int(dataframe[dataframe['Código'] == int(ug)]['Energia Injetada PT'].values[0])
  consumo_ug = int(dataframe[dataframe['Código'] == int(ug)]['Energia Ativa FP'].values[0]) + int(dataframe[dataframe['Código'] == int(ug)]['Energia Ativa PT'].values[0])
  carteira_anterior = int(dataframe[dataframe['Código'] == int(ug)]['Saldo Anterior FP'].values[0]) + int(dataframe[dataframe['Código'] == int(ug)]['Saldo Anterior PT'].values[0])
  carteira_atual = int(dataframe[dataframe['Código'] == int(ug)]['Saldo Final FP'].values[0]) + int(dataframe[dataframe['Código'] == int(ug)]['Saldo Final PT'].values[0])
  transferencia_ug = int(dataframe[dataframe['Código'] == int(ug)]['Saldo Transferido Outra UC FP'].values[0]) + int(dataframe[dataframe['Código'] == int(ug)]['Saldo Transferido Outra UC PT'].values[0])

  if consumo_calc == consumo_ug and geracao_calc == geracao_ug and transferido_calc == transferencia_ug:

    print('                  Medições Digitadas Corretas!!! \n')
    print('Consumo Digitado: {}'.format(consumo_calc))
    print('Consumo Real: {} \n'.format(consumo_ug))

    print('Geração Digitado: {}'.format(geracao_calc))
    print('Geração Real: {} \n'.format(geracao_ug))

    print('Transferência Digitada: {}'.format((transferido_calc)))
    print('Transferência Real: {} \n'.format(transferencia_ug))
    print('Transferência para a Carteira da UG: {}'.format(carteira_atual - carteira_anterior))

    return True
  
  else:
  
    print('                  Medições Digitadas Incorretas!!! \n')
    print('Consumo Digitado: {}'.format(consumo_calc))
    print('Consumo Real: {}'.format(consumo_ug))

    print('Geração Digitado: {}'.format(geracao_calc))
    print('Geração Real: {}'.format(geracao_ug))

    print('Transferência Digitada: {}'.format((transferido_calc)))
    print('Transferência Real: {}'.format(transferencia_ug))

    return False

def validacao_recebimentos(df_comparativo):
  
  global transferencia_ug
  global dataframe_divergencias

  somatorio_recebimentos = df_comparativo['Recebimento AXS'].sum()
  somatorio_distribuidora = df_comparativo['Recebimento Distribuidora'].sum() + df_comparativo['Diferenca'].sum()

  try:
    if (somatorio_recebimentos - int(transferencia_ug))/int(transferencia_ug)*100 < 1 and (somatorio_distribuidora - int(transferencia_ug))/int(transferencia_ug)*100 < 1:
      print("Sucesso na conferência:  Energia Transferida->{} = Recebimentos AXS->{} = Recebimentos Distribuidora->{} \n".format(int(transferencia_ug), somatorio_recebimentos, somatorio_distribuidora, ))
      
      dataframe_divergencias = df_comparativo[(df_comparativo['Diferenca'] > 2) | (df_comparativo['Diferenca'] < -2)]
      divergencias = dataframe_divergencias.shape[0]

      if divergencias == 0:
        return True
      else:
        False
    else:
      print("Iconsistência na conferência: {} = {} = {}".format(int(transferencia_ug), somatorio_distribuidora, somatorio_recebimentos))
      return False
  
  except ZeroDivisionError:
     print("Error: Cannot divide by zero!")

def save_csv(df_a, df_r, df_s):

  global distribuidora
  global competencia
  global ug
  global arquivo_aberto

  diretorio = 'C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento/Relatórios de Recebimentos/{}/{}/{}'.format(distribuidora, competencia, ug)
  if not os.path.exists(diretorio):
    os.makedirs(diretorio)
    with pandas.ExcelWriter(os.path.join(diretorio, '{}c.xlsx'.format(ug)), engine='xlsxwriter') as writer:
      df_a.to_excel(writer, sheet_name='Aprovados', index=False)
      df_r.to_excel(writer, sheet_name='Reprovados', index=False)
      df_s.to_excel(writer, sheet_name='Sem Compensação', index=False)

  origem = caminho_pdf
  destino = os.path.join(diretorio, 'DemonstrativoMicroMiniGeracao.pdf')

  arquivo_aberto.close()
  shutil.move(origem, destino)

def csv(data: datetime, ug: str, df: pandas.DataFrame):

    data_hoje = datetime.today()
    data_hoje = data_hoje.strftime('%Y.%m.%d')

    dataframe = pandas.DataFrame(columns=['Periodo', 'Modalidade', 'Instalacao', 'Quota', 'Posto_horario', 'Qtd_consumo', 'Qtd_geracao', 'Qtd_compensacao', 'Qtd_saldo_ant', 'Qtd_transferencia', 'Qtd_recebimento', 'Qtd_saldo_atual', 'Qtd_saldo_exp', 'Qtd_prox_saldo_exp', 'Per_prox_saldo_exp'])

    for index, row in df.iterrows():
        if index == 0:
            Periodo = f'{dicio_ano[data[3:5]]}/{dicio_meses[data[0:3]]}'
            Modalidade = "Ger. Compartilhada-Geradora"
            Instalacao = ug
            Quota = 0
            Posto_horario = "FORAP"
            Qtd_consumo = ""
            Qtd_geracao = ""
            Qtd_compensacao = ""
            Qtd_saldo_ant = ""
            Qtd_transferencia = ""
            Qtd_recebimento = ""
            Qtd_saldo_atual = ""
            Qtd_saldo_exp = ""
            Qtd_prox_saldo_exp = ""
            Per_prox_saldo_exp = ""
        elif index == len(df) - 1:
            Periodo = f'{dicio_ano[data[3:5]]}/{dicio_meses[data[0:3]]}'
            Modalidade = "Ger. Compartilhada-Geradora"
            Instalacao = ug
            Quota = 0
            Posto_horario = "FORAP"
            Qtd_consumo = int(consumo_calc)
            Qtd_geracao = int(geracao_calc)
            Qtd_compensacao = compensado_calc
            Qtd_saldo_ant = ""
            Qtd_transferencia = transferido_calc
            Qtd_recebimento = ""
            Qtd_saldo_atual = ""
            Qtd_saldo_exp = ""
            Qtd_prox_saldo_exp = ""
            Per_prox_saldo_exp = ""
        else:
            Periodo = f'{dicio_ano[data[3:5]]}/{dicio_meses[data[0:3]]}'
            Modalidade = "Ger. Compartilhada-Recebedora"
            Instalacao = row["Código"]
            Quota = row["Quota"]
            Posto_horario = "FORAP"
            Qtd_consumo = ""
            Qtd_geracao = ""
            Qtd_compensacao = ""
            Qtd_saldo_ant = ""
            Qtd_transferencia = ""
            Qtd_recebimento = int(row["Recebimento Distribuidora"])
            Qtd_saldo_atual = ""
            Qtd_saldo_exp = ""
            Qtd_prox_saldo_exp = ""
            Per_prox_saldo_exp = ""

        dataframe.loc[len(dataframe)] = [Periodo, Modalidade, Instalacao, Quota, Posto_horario, Qtd_consumo, Qtd_geracao, Qtd_compensacao, Qtd_saldo_ant, Qtd_transferencia, Qtd_recebimento, Qtd_saldo_atual, Qtd_saldo_exp, Qtd_prox_saldo_exp, Per_prox_saldo_exp]

    dataframe['Instalacao'] = dataframe['Instalacao'].astype(int)
    dataframe["Quota"] = dataframe["Quota"].astype(float) / 100

    nome_arquivo = '{} {}.csv'.format(data_hoje, ug)

    caminho_arquivo = os.path.join('C:/Users/lucas.conceicao/Downloads', nome_arquivo)

    dataframe.to_csv(caminho_arquivo, index=False, sep=";", encoding="utf-8-sig", decimal=',')
    print(dataframe)

    print(f"Arquivo CSV {os.path.basename(caminho_arquivo)} criado com sucesso!")

def recusados ():

    pergunta = input("Deseja realizar a função de adicionar os recusados no arquivo de conferência na pasta de destino: ")
    
    if pergunta.lower() == "s":
       
      try:
          workbook = load_workbook(caminho_recusados)
          sheet = workbook[workbook.sheetnames[0]]
          
          df_temp = dataframe_recusados[['Comp', 'Distribuidora', 'Usina', 'UG', 'UC', 'Motivo', 'Energia']].copy()

          star_row = sheet.max_row + 1
          n_linhas = df_temp.shape[0]

          sheet.insert_rows(star_row, amount=n_linhas)

          current_row = star_row
          for idx, row in df_temp.iterrows():
            sheet.cell(row=current_row, column=1, value=row["Comp"])
            sheet.cell(row=current_row, column=2, value=row["Distribuidora"])
            sheet.cell(row=current_row, column=3, value=row["Usina"])
            sheet.cell(row=current_row, column=4, value=row["UG"])
            sheet.cell(row=current_row, column=5, value=row["UC"])
            sheet.cell(row=current_row, column=6, value=row["Motivo"])
            sheet.cell(row=current_row, column=7, value=row["Energia"])
            current_row += 1

          workbook.save(caminho_recusados)
          print("Clientes recusados foram adicionados no arquivo de referência")

      except FileNotFoundError:
          print(f"O arquivo não foi encontrato.")
    else:
      print("Não há recusados")

def nao_compensadoss ():

    pergunta = input("Deseja realizar a função de adicionar os sem compensação no arquivo de conferência na pasta de destino: ")
    
    if pergunta.lower() == "s":
       
      nao_compensados = []
      try:
          workbook = load_workbook(caminho_compensados)
          sheet = workbook['Planilha1']

          for row in sem_compensacao.itertuples():
              nao_compensados.append(row[1])
              nao_compensados.append(ug)
              nao_compensados.append(row[2])

              sheet.append(nao_compensados)
              workbook.save(caminho_compensados)

              nao_compensados = []

              print("Clientes sem compensação foram adicionados no arquivo de referência")

      except FileNotFoundError:
          print(f"O arquivo não foi encontrato.")
      except Exception as e:
          print(f"Ocorreu um erro: {e}")
    else:
      print("Não há clientes que não foram compensados")

def relatorio_de_compensacao_organizador(data):
  
  global geracao_ug
  global consumo_ug
  global transferencia_ug
  global dataframe_recusados
  global sem_compensacao
  global percentual_recusados

  dataframe = pandas.DataFrame(columns=['Código', 
                                        'Referência',
                                        'Saldo Anterior PT',
                                        'Saldo Anterior FP',
                                        'Saldo Anterior TP',
                                        'Cred. Receb. Outra UC PT',
                                        'Cred. Receb. Outra UC FP',
                                        'Cred. Receb. Outra UC TP',
                                        'Energia Injetada PT',
                                        'Energia Injetada FP',
                                        'Energia Injetada TP',
                                        'Energia Ativa PT',
                                        'Energia Ativa FP',
                                        'Energia Ativa TP',
                                        'Credito Utilizado no Mês PT',
                                        'Credito Utilizado no Mês FP',
                                        'Credito Utilizado no Mês TP',
                                        'Percentual',
                                        'Prioridade',
                                        'Saldo Transferido Outra UC PT',
                                        'Saldo Transferido Outra UC FP',
                                        'Saldo Transferido Outra UC TP',
                                        'Saldo Final PT',
                                        'Saldo Final FP',
                                        'Saldo Final TP'])
  
  arquivo_lido, quantidade_paginas = leitor_arquivo()

  for pagina in range(quantidade_paginas):
    if check_empty(pagina) or check_columns(pagina) == 10:
      pass
    else:
      if pagina == 0:
        codigo_uc = unidade_geradora(pagina, arquivo_lido)
        uc_dataframe = dataframe_pandas_copel(codigo_uc, pagina)
      else:
        codigo_uc = unidade_consumidora(pagina, arquivo_lido)
        uc_dataframe = dataframe_pandas_copel(codigo_uc, pagina)
    
    dataframe = pandas.concat([dataframe, uc_dataframe], axis=0)
    dataframe['Código'] = dataframe['Código'].astype(int)

  percentual_recusados = dataframe[dataframe['Código'] == int(ug)]['Percentual'].values[0]

  if len(percentual_recusados) >= 4:
        percentual_recusados = float(percentual_recusados)

  elif percentual_recusados[0] == str(0):
      percentual_recusados = float(percentual_recusados)

  elif percentual_recusados[0] == ".":
    percentual_recusados = float(f'0{percentual_recusados}')

  else:
     print("Erro ao ler o percentual de recusados")

  dataframe_analise_de_faturas(data)

  excedente = excedente_geracao(data)
  
  dataframe_recebimento = dataframe_pandas_recebimento(data, excedente)

  dataframe_recebimento['UC'] = dataframe_recebimento['UC'].astype(int)

  if validacao_medicoes(dataframe):

    dataframe_COPEL = dataframe[(dataframe['Referência'] == "10/2026") & (dataframe['Cred. Receb. Outra UC FP'] != "0")]
    sem_compensacao = dataframe_COPEL[(dataframe_COPEL['Energia Ativa TP'] != "0") & (dataframe_COPEL['Credito Utilizado no Mês TP'] == "0")]
    sem_compensacao = sem_compensacao[['Referência', 'Código']]
    sem_recebimento = dataframe[(dataframe['Referência'] == "10/2026") & (dataframe['Cred. Receb. Outra UC FP'] == "0")]
    sem_recebimento = sem_recebimento[['Referência', 'Código']]

    dataframe_aprovados = dataframe_pandas_comparativo(dataframe_COPEL, dataframe_recebimento)
  
    if validacao_recebimentos(dataframe_aprovados):
      print("                     Todos os créditos recebidos pelos beneficiários estão corretos \n ")
      print('                                        RESUMO DOS BENEFICIÁRIOS \n')
      print('                                ************  APROVADOS  ************ \n')
      colunas = dataframe_aprovados.columns.tolist()
      print(tabulate(dataframe_aprovados, headers=colunas, tablefmt='fancy_grid'))
      print(' ')
      print('                                 ***********  RECUSADOS  ************')
      colunas = dataframe_recusados.columns.tolist()
      print(tabulate(dataframe_recusados, headers=colunas, tablefmt='fancy_grid'))
      print(' ')
      print('                              ***********  SEM COMPENSAÇÃO  ************\n')
      colunas = sem_compensacao.columns.tolist()
      print(tabulate(sem_compensacao, headers=colunas, tablefmt='fancy_grid'))
      print(' ')
      #print('                              ***********  SEM RECEBIMENTOS  ************\n')
      #colunas = sem_recebimento.columns.tolist()
      #print(tabulate(sem_recebimento, headers=colunas, tablefmt='fancy_grid'))

      if len(dataframe_recusados) != 0:
          recusados()
    
      if len(sem_compensacao) != 0:
          nao_compensadoss()

      pergunta = input("Deseja realizar a função de criação e salvamento do arquivo de conferência na pasta de destino: ")

      if pergunta.lower() == "s":
        save_csv(dataframe_aprovados, dataframe_recusados, sem_compensacao)
        csv(data, ug, dataframe_aprovados)
      else:
          print("Não foi criado o arquivo de conferência")
    
    else:
        print('                               Há divergências de alguns créditos \n \n ')
        print('                                    RESUMO DOS BENEFICIÁRIOS \n \n ')
        colunas = dataframe_divergencias.columns.tolist()
        print(tabulate(dataframe_divergencias, headers=colunas, tablefmt='fancy_grid'))
        print("\nNão foi criado o arquivo de conferência")
        dataframe_divergencias.to_excel(f'divergentes.xlsx')


relatorio_de_compensacao_organizador('OUT26')
