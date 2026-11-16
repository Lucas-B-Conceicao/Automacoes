from tabula import read_pdf
import PyPDF2
import pandas
import regex
import os
import shutil
from openpyxl import load_workbook
from tabulate import tabulate

caminho_pdf = 'C:/Users/lucas.conceicao/Downloads/DemonstrativoMicroMiniGeracao.pdf'
caminho_recebimento = 'C:/Users/lucas.conceicao/AXS ENERGIA S A/Usinas AXS - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento'
caminho_recusados = "C:/Users/lucas.conceicao/AXS ENERGIA S A/Usinas AXS - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento/Extras/Recusados.xlsx"
caminho_compensados = "C:/Users/lucas.conceicao/AXS ENERGIA S A/Usinas AXS - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento/Extras/Sem_compensacao.xlsx"
distribuidoras = {'CEMIG': ['3014629309', '3014629582', '3014629614', '3014629370', '3014109500', '3014109501', '3014109502', '3014233708', '3015083428', '3015083492', '3014926608', '3014926612', '3014233705', '3014298175', '3014298176', '3014427046', '3015120874', '3014292343', '3014292356', '3014292382', '3014837599', '3014837605', '3014837606', '3014902567', '3014902579', '3014902584', '3014900389', '3014900403', '3014900405', '3014868833', '3014868836','3014868860'], 
                  'COPEL': ['110069781','110069749','110069692','110022920','110022890','110635710','110519370','110519256','110519329','110635809', '108883663', '108883116', '108883590', '108288897', '108288854', '108288927', '108289010', '108288633', '110510526', '110510429', '110510640', '110756517', '110756428', '110756606', '110411889', '110411978', '110411854', '110877306', '110877519', '110877446', '110877845', '110877918', '110878019', '110876806', '110876865', '110876393', '110517997', '110517890', '110518071', '110837150', '109368630', '110509730', '110509838'], 
                  'CPFL': ['4003494881', '4003310119', '4003427926', '4003518327', '4003591357', '4003766488', '4003458619', '4003458619', '4003700263', '4003700370', '4003510512', '4003459206', '4003549082'], 
                  'ELEKTRO': ['45709033', '45709254', '45709360', '45963592', '45963894', '45948194', '45948160', '45948321'], 
                  'ENERGISA': ['4200568', '4200527', '4200398', '4243410', '4243513', '4243527', '4518372', '4518402', '4518482', '4400952', '4401014', '4401020', '4385377', '4385297', '4385076']}
recusadas = []
energia_recusadas = []
distribuidora = None
competencia = None
ug = None

geracao_ug = None
consumo_ug = None
transferencia_ug = None

geracao_calc = None
consumo_calc = None
transferido_calc = None

dataframe_aprovados = None
dataframe_recusados = None
sem_compensacao = None
dataframe_divergencias = None

def leitor_arquivo():

  global caminho_pdf

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

  if df.shape[1] == 26:
    df.rename(columns={'': 'Indice',
                    'Referência': 'Referência',
                    'Saldo Anterior': 'Saldo Anterior PT', 'Cred. Receb. Outra UC': 'Saldo Anterior FP', 'Energia Injetada': 'Saldo Anterior TP', 
                    'Energia Ativa': 'Cred. Receb. Outra UC PT', 'Credito Utilizado no Mês': 'Cred. Receb. Outra UC FP', 'Saldo Mes': 'Cred. Receb. Outra UC TP', 
                    'Saldo Transferido Outra UC': 'Energia Injetada PT', 'Saldo Final': 'Energia Injetada FP', 'Unnamed: 0': 'Energia Injetada TP', 
                    'Unnamed: 1': 'Energia Ativa PT', 'Unnamed: 2': 'Energia Ativa FP', 'Unnamed: 3': 'Energia Ativa TP', 
                    'Unnamed: 4': 'Credito Utilizado no Mês PT', 'Unnamed: 5': 'Credito Utilizado no Mês FP', 'Unnamed: 6': 'Credito Utilizado no Mês TP',
                    'Unnamed: 7': 'Saldo Mes PT', 'Unnamed: 8': 'Saldo Mes FP', 'Unnamed: 9': 'Saldo Mes TP',
                    'Unnamed: 10': 'Saldo Transferido Outra UC PT', 'Unnamed: 11': 'Saldo Transferido Outra UC FP', 'Unnamed: 12': 'Saldo Transferido Outra UC TP',
                    'Unnamed: 13': 'Saldo Final PT', 'Unnamed: 14': 'Saldo Final FP', 'Unnamed: 15': 'Saldo Final TP'},
                        inplace=True)
    
    dados_coluna_uc = [uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc]
    df.insert(1, 'Código', dados_coluna_uc)

    df.drop(df.head(1).index, inplace=True)
    df.drop(df.columns[0], axis=1)

  if df.shape[1] == 18:
    if pagina == 0:
        df.rename(columns={'': 'Indice',
                            'Referência': 'Referência',
                            'Saldo Anterior': 'Saldo Anterior PT', 'Cred. Receb. Outra UC': 'Saldo Anterior FP',
                            'Energia Injetada': 'Cred. Receb. Outra UC PT', 'Energia Ativa': 'Cred. Receb. Outra UC FP',
                            'Credito Utilizado no Mês': 'Energia Injetada PT', 'Saldo Mes': 'Energia Injetada FP',
                            'Saldo Transferido Outra UC': 'Energia Ativa PT', 'Saldo Final': 'Energia Ativa FP',
                            'Unnamed: 0': 'Credito Utilizado no Mês PT', 'Unnamed: 1': 'Credito Utilizado no Mês FP',
                            'Unnamed: 2': 'Saldo Mes PT', 'Unnamed: 3': 'Saldo Mes FP',
                            'Unnamed: 4': 'Saldo Transferido Outra UC PT', 'Unnamed: 5': 'Saldo Transferido Outra UC FP',
                            'Unnamed: 6': 'Saldo Final PT', 'Unnamed: 7': 'Saldo Final FP'},
                        inplace=True)
        
        dados_nova_coluna = [0,0,0,0,0,0,0,0,0,0,0,0,0,0]
        df.insert(4, 'Saldo Anterior TP', dados_nova_coluna)
        df.insert(7, 'Cred. Receb. Outra UC TP', dados_nova_coluna)
        df.insert(10, 'Energia Injetada TP', dados_nova_coluna)
        df.insert(13, 'Energia Ativa TP', dados_nova_coluna)
        df.insert(16, 'Credito Utilizado no Mês TP', dados_nova_coluna)
        df.insert(19, 'Saldo Mes TP', dados_nova_coluna)
        df.insert(22, 'Saldo Transferido Outra UC TP', dados_nova_coluna)
        df.insert(25, 'Saldo Final TP', dados_nova_coluna)
    else:
        df.rename(columns={'': 'Indice',
                            'Referência': 'Referência',
                            'Saldo Anterior': 'Saldo Anterior FP', 'Cred. Receb. Outra UC': 'Saldo Anterior TP',
                            'Energia Injetada': 'Cred. Receb. Outra UC FP', 'Energia Ativa': 'Cred. Receb. Outra UC TP',
                            'Credito Utilizado no Mês': 'Energia Injetada FP', 'Saldo Mes': 'Energia Injetada TP',
                            'Saldo Transferido Outra UC': 'Energia Ativa FP', 'Saldo Final': 'Energia Ativa TP',
                            'Unnamed: 0': 'Credito Utilizado no Mês FP', 'Unnamed: 1': 'Credito Utilizado no Mês TP',
                            'Unnamed: 2': 'Saldo Mes FP', 'Unnamed: 3': 'Saldo Mes TP',
                            'Unnamed: 4': 'Saldo Transferido Outra UC FP', 'Unnamed: 5': 'Saldo Transferido Outra UC TP',
                            'Unnamed: 6': 'Saldo Final FP', 'Unnamed: 7': 'Saldo Final TP'},
                        inplace=True)
        
        dados_nova_coluna = [0,0,0,0,0,0,0,0,0,0,0,0,0,0]
        df.insert(2, 'Saldo Anterior PT', dados_nova_coluna)
        df.insert(5, 'Cred. Receb. Outra UC PT', dados_nova_coluna)
        df.insert(8, 'Energia Injetada PT', dados_nova_coluna)
        df.insert(11, 'Energia Ativa PT', dados_nova_coluna)
        df.insert(14, 'Credito Utilizado no Mês PT', dados_nova_coluna)
        df.insert(17, 'Saldo Mes PT', dados_nova_coluna)
        df.insert(20, 'Saldo Transferido Outra UC PT', dados_nova_coluna)
        df.insert(23, 'Saldo Final PT', dados_nova_coluna)

        dados_coluna_uc = [uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc, uc]
        df.insert(1, 'Código', dados_coluna_uc)

        df.drop(df.head(1).index, inplace=True)
        df.drop(df.columns[0], axis=1)

  os.remove('{}.csv'.format(uc))

  return df

def dataframe_pandas_recebimento(data):
  
  global caminho_recebimento
  global ug
  global geracao_calc
  global consumo_calc
  global transferido_calc
  
  arquivo = "{} - Todas.xlsm".format(data)
  caminho_recebimento_arq = os.path.join(caminho_recebimento, arquivo)
  dataframe = pandas.read_excel(caminho_recebimento_arq, ug)

  consumo_calc = int(dataframe[dataframe['Instalacao'] == int(ug)]['Qtd_consumo'].sum())
  geracao_calc = int(dataframe[dataframe['Instalacao'] == int(ug)]['Qtd_geracao'].sum())
  transferido_calc = int(dataframe[dataframe['Instalacao'] == int(ug)]['Qtd_transferencia'].sum())

  dataframe = dataframe.drop(index=0)
  dataframe = dataframe.drop(index=len(dataframe))

  return dataframe

def dataframe_pandas_comparativo(df_copel, df_recebimento):
  
  global recusadas
  global energia_recusadas
  global dataframe_recusados

  dataframe_aprovados = pandas.DataFrame(columns=['Código', 'Recebimento AXS',
                                        'Recebimento Distribuidora', 'Diferenca', 'variacao'])

  for index, row in df_recebimento.iterrows():
    if index != 0 and index != len(dataframe_aprovados)-1 and index != len(dataframe_aprovados):
        if str(row['Instalacao']) in df_copel['Código'].values:
            dataframe_aprovados.loc[len(dataframe_aprovados)] = [str(row['Instalacao']), int(row['Qtd_recebimento']), int(df_copel[df_copel['Código'] == str(row['Instalacao'])]['Cred. Receb. Outra UC FP'].values[0]) + int(df_copel[df_copel['Código'] == str(row['Instalacao'])]['Cred. Receb. Outra UC PT'].values[0]), int(row['Qtd_recebimento'])-(int(df_copel[df_copel['Código'] == str(row['Instalacao'])]['Cred. Receb. Outra UC FP'].values[0])+int(df_copel[df_copel['Código'] == str(row['Instalacao'])]['Cred. Receb. Outra UC PT'].values[0])), (int(row['Qtd_recebimento'])-(int(df_copel[df_copel['Código'] == str(row['Instalacao'])]['Cred. Receb. Outra UC FP'].values[0])+int(df_copel[df_copel['Código'] == str(row['Instalacao'])]['Cred. Receb. Outra UC PT'].values[0])))/(int(df_copel[df_copel['Código'] == str(row['Instalacao'])]['Cred. Receb. Outra UC FP'].values[0])+int(df_copel[df_copel['Código'] == str(row['Instalacao'])]['Cred. Receb. Outra UC PT'].values[0]))*100]
        else:
            recusadas.append(row['Instalacao'])
            energia_recusadas.append(row['Qtd_recebimento'])
  
  dataframe_recusados = pandas.DataFrame({'Código': recusadas, 'Recebimento AXS': energia_recusadas})
  #print(list(zip(recusadas, energia_recusadas)))
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

  if consumo_calc == consumo_ug and geracao_calc == geracao_ug: #and transferido_calc - (carteira_atual - carteira_anterior) == transferencia_ug:

    print('                  Medições Digitadas Corretas!!! \n')
    print('Consumo Digitado: {}'.format(consumo_calc))
    print('Consumo Real: {} \n'.format(consumo_ug))

    print('Geração Digitado: {}'.format(geracao_calc))
    print('Geração Real: {} \n'.format(geracao_ug))

    print('Transferência Digitada: {}'.format((transferido_calc - (carteira_atual - carteira_anterior))))
    print('Transferência Real: {} \n'.format(transferencia_ug))
    print('Transferência para a Carteira da UG: {}'.format(carteira_atual - carteira_anterior))

    return True
  
  else:
  
    print('                  Medições Digitadas Incorretas!!! \n')
    print('Consumo Digitado: {}'.format(consumo_calc))
    print('Consumo Real: {}'.format(consumo_ug))

    print('Geração Digitado: {}'.format(geracao_calc))
    print('Geração Real: {}'.format(geracao_ug))

    print('Transferência Digitada: {}'.format((transferido_calc - (carteira_atual - carteira_anterior))))
    print('Transferência Real: {}'.format(transferencia_ug))

    return False

def validacao_recebimentos(df_comparativo):
  
  global transferencia_ug
  global dataframe_divergencias

  somatorio_recebimentos = df_comparativo['Recebimento AXS'].sum()
  somatorio_distribuidora = df_comparativo['Recebimento Distribuidora'].sum() + df_comparativo['Diferenca'].sum()

  try:
    if (somatorio_recebimentos - int(transferencia_ug))/int(transferencia_ug)*100 < 1 and (somatorio_distribuidora - int(transferencia_ug))/int(transferencia_ug)*100 < 1:
      print("Sucesso na conferência:  Energia Transferida->{} = Recebimentos AXS->{} = Recebimentos Distribuidora->{} \n".format(int(transferencia_ug), somatorio_distribuidora, somatorio_recebimentos))
      
      dataframe_divergencias = df_comparativo[(df_comparativo['Diferenca'] > 1) | (df_comparativo['Diferenca'] < -1)]
      divergencias = dataframe_divergencias.shape[0]

      if divergencias == 0:
        return True
      else:
        False
    else:
      print("Inconsistência na conferência: {} = {} = {}".format(int(transferencia_ug), somatorio_distribuidora, somatorio_recebimentos))
      return False
  
  except ZeroDivisionError:
     print("Error: Cannot divide by zero!")

def save_csv(df_a, df_r, df_s):

  global distribuidora
  global competencia
  global ug

  diretorio = 'C:/Users/lucas.conceicao/AXS ENERGIA S A/Usinas AXS - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento/Relatórios de Recebimentos/{}/{}/{}'.format(distribuidora, competencia, ug)
  if not os.path.exists(diretorio):
    os.makedirs(diretorio)
    with pandas.ExcelWriter(os.path.join(diretorio, '{}c.xlsx'.format(ug)), engine='xlsxwriter') as writer:
      df_a.to_excel(writer, sheet_name='Aprovados', index=False)
      df_r.to_excel(writer, sheet_name='Reprovados', index=False)
      df_s.to_excel(writer, sheet_name='Sem Compensação', index=False)

  origem = caminho_pdf
  destino = os.path.join(diretorio, 'DemonstrativoMicroMiniGeracao.pdf')
  shutil.move(origem, destino)

def recusados ():

    recusado = []
    try:
        workbook = load_workbook(caminho_recusados)
        sheet = workbook['Historico']

        for row in dataframe_recusados.itertuples():
            recusado.append(competencia)
            recusado.append(distribuidora)
            recusado.append('UFV_NOME')
            recusado.append(ug)
            recusado.append(row[1])
            recusado.append('Sem informação')
            recusado.append(row[2])

            sheet.append(recusado)
            workbook.save(caminho_recusados)

            recusado = []

            print("Clientes recusados foram adicionados no arquivo de referência")

    except FileNotFoundError:
        print(f"O arquivo não foi encontrato.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

def nao_compensadoss ():

    nao_compensados = []
    try:
        workbook = load_workbook(caminho_compensados)
        sheet = workbook['Planilha1']

        for row in sem_compensacao.itertuples():
            nao_compensados.append(row[2])
            nao_compensados.append(ug)
            nao_compensados.append(row[1])

            sheet.append(nao_compensados)
            workbook.save(caminho_compensados)

            nao_compensados = []

            print("Clientes sem compensação foram adicionados no arquivo de referência")

    except FileNotFoundError:
        print(f"O arquivo não foi encontrato.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

def relatorio_de_compensacao_organizador(data):
  
  global geracao_ug
  global consumo_ug
  global transferencia_ug
  global dataframe_recusados
  global sem_compensacao

  dataframe = pandas.DataFrame(columns=['Código', 'Referência',
                                        'Saldo Anterior PT','Saldo Anterior FP','Saldo Anterior TP',
                                        'Cred. Receb. Outra UC PT','Cred. Receb. Outra UC FP','Cred. Receb. Outra UC TP',
                                        'Energia Injetada PT','Energia Injetada FP','Energia Injetada TP',
                                        'Energia Ativa PT','Energia Ativa FP','Energia Ativa TP',
                                        'Credito Utilizado no Mês PT','Credito Utilizado no Mês FP','Credito Utilizado no Mês TP',
                                        'Saldo Mes PT','Saldo Mes FP','Saldo Mes TP',
                                        'Saldo Transferido Outra UC PT','Saldo Transferido Outra UC FP','Saldo Transferido Outra UC TP',
                                        'Saldo Final PT','Saldo Final FP','Saldo Final TP'])
  
  arquivo_lido, quantidade_paginas = leitor_arquivo()

  for pagina in range(quantidade_paginas-192):
    if pagina == 0:
      codigo_uc = unidade_geradora(pagina, arquivo_lido)
      uc_dataframe = dataframe_pandas_copel(codigo_uc, pagina)
    else:
      codigo_uc = unidade_consumidora(pagina, arquivo_lido)
      uc_dataframe = dataframe_pandas_copel(codigo_uc, pagina)

    dataframe = pandas.concat([dataframe, uc_dataframe], axis=0)

  dataframe_recebimento = dataframe_pandas_recebimento(data)

  if validacao_medicoes(dataframe):

    dataframe_COPEL = dataframe[(dataframe['Referência'] == "12/2025") & (dataframe['Cred. Receb. Outra UC FP'] != "0")]
    sem_compensacao = dataframe_COPEL[(dataframe_COPEL['Energia Ativa TP'] != "0") & (dataframe_COPEL['Credito Utilizado no Mês TP'] == "0")]
    sem_recebimento = dataframe[(dataframe['Referência'] == "12/2025") & (dataframe['Cred. Receb. Outra UC FP'] == "0")]

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
      print('                              ***********  SEM RECEBIMENTOS  ************\n')
      colunas = sem_recebimento.columns.tolist()
      print(tabulate(sem_recebimento, headers=colunas, tablefmt='fancy_grid'))

      if len(dataframe_recusados) != 0:
        pergunta = input("Deseja realizar a função de adicionar os recusados no arquivo de conferência na pasta de destino: ")

        if pergunta.lower() == "s":
            recusados()
        else:
            print("Não há recusados")
    
      if len(sem_compensacao) != 0:
        pergunta = input("Deseja realizar a função de adicionar os sem compensação no arquivo de conferência na pasta de destino: ")

        if pergunta.lower() == "s":
            nao_compensadoss()
        else:
            print("Não há clientes que não foram compensados")

      pergunta = input("Deseja realizar a função de criação e salvamento do arquivo de conferência na pasta de destino: ")

      if pergunta.lower() == "s":
        save_csv(dataframe_aprovados, dataframe_recusados, sem_compensacao)
      else:
          print("Não foi criado o arquivo de conferência")
    
    else:
        print('                               Há divergências de alguns créditos \n \n ')
        print('                                    RESUMO DOS BENEFICIÁRIOS \n \n ')
        colunas = dataframe_divergencias.columns.tolist()
        print(tabulate(dataframe_divergencias, headers=colunas, tablefmt='fancy_grid'))
        print("\nNão foi criado o arquivo de conferência")

relatorio_de_compensacao_organizador('DEZ25')
