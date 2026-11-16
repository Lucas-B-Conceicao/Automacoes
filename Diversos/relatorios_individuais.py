from tabula import read_pdf
import PyPDF2
import pandas
import regex
import os

caminho_pdf = 'C:/Users/lucas.conceicao/Downloads/DemonstrativoMicroMiniGeracao.pdf'
distribuidoras = distribuidoras = {'CEMIG': ['3014629309', '3014629582', '3014629614', '3014629370', '3014109500', '3014109501', '3014109502', '3014233708', '3015083428', '3015083492', '3014926608', '3014926612', '3014233705', '3014298175', '3014298176', '3014427046', '3015120874', '3014292343', '3014292356', '3014292382', '3014837599', '3014837605', '3014837606', '3014902567', '3014902579', '3014902584', '3014900389', '3014900403', '3014900405', '3014868833', '3014868836','3014868860'], 
                  'COPEL': ['110635710','110519370','110519256','110519329','110635809', '108883663', '108883116', '108883590', '108288897', '108288854', '108288927', '108289010', '108288633', '110510526', '110510429', '110510640', '110756517', '110756428', '110756606', '110411889', '110411978', '110411854', '110877306', '110877519', '110877446', '110877845', '110877918', '110878019', '110876806', '110876865', '110876393', '110517997', '110517890', '110518071', '110837150', '109368630', '110509730', '110509838'], 
                  'CPFL': ['4003494881', '4003310119', '4003427926', '4003518327', '4003591357', '4003766488', '4003458619', '4003458619', '4003700263', '4003700370', '4003510512', '4003459206', '4003549082'], 
                  'ELEKTRO': ['45709033', '45709254', '45709360', '45963592', '45963894', '45948194', '45948160', '45948321'], 
                  'ENERGISA': ['4200568', '4200527', '4200398', '4243410', '4243513', '4243527', '4518372', '4518402', '4518482', '4400952', '4401014', '4401020', '4385377', '4385297', '4385076']}
recusadas = []
distribuidora = None
competencia = None
ug = None

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
  ug = encontrar_uc = regex.findall(r'UC\s*:\s(\d+)', conteudo)[0]
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
  encontrar_uc = regex.findall(r'UC\s*:\s(\d+)', conteudo)[0]

  tabelas = read_pdf(caminho_pdf, pages=pagina+1, multiple_tables=True, lattice=True, pandas_options={'header': 0})
  tabelas = tabelas[1]
  tabelas.to_csv('{}.csv'.format(encontrar_uc), encoding='utf-8')

  return encontrar_uc

def dataframe_pandas(uc, pagina):

  global distribuidora
  global competencia
  global ug

  diretorio = 'C:/Users/lucas.conceicao/Relatorio de Compensacao/{}/{}/{}'.format(distribuidora, competencia, ug)

  if not os.path.exists(diretorio):
    os.makedirs(diretorio)

  if os.path.isfile(os.path.join(diretorio, '{}.csv'.format(uc))):
    print('O arquivo {}.csv já existe no local'.format(uc))
  else:
    df = pandas.read_csv('{}.csv'.format(uc), sep=',')
    print('O arquivo {}.csv foi criado'.format(uc))

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

    df.drop(df.head(1).index, inplace=True)
    df.drop(df.columns[0], axis=1)

    df.to_csv(os.path.join(diretorio, '{}.csv'.format(uc)), sep=',', index=False, encoding='utf-8')
  os.remove('{}.csv'.format(uc))

def relatorio_de_compensacao_organizador():

  arquivo_lido, quantidade_paginas = leitor_arquivo()

  for pagina in range(quantidade_paginas):
    if pagina == 0:
      codigo_uc = unidade_geradora(pagina, arquivo_lido)
      dataframe_pandas(codigo_uc, pagina)
    else:
      codigo_uc = unidade_consumidora(pagina, arquivo_lido)
      dataframe_pandas(codigo_uc, pagina)


relatorio_de_compensacao_organizador()