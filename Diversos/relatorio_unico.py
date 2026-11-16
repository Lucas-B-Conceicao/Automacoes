from tabula import read_pdf
import PyPDF2
import pandas
import regex
import os

caminho_pdf = 'C:/Users/lucas.conceicao/Downloads/DemonstrativoMicroMiniGeracao.pdf'
distribuidoras = { 
                  'COPEL': ['111254310','110878957','110635710','110519370','110519256','110519329','110635809', '108883663', '108883116', '108883590', '108288897', '108288854', '108288927', '108289010', '108288633', '110510526', '110510429', '110510640', '110756517', '110756428', '110756606', '110411889', '110411978', '110411854', '110877306', '110877519', '110877446', '110877845', '110877918', '110878019', '110876806', '110876865', '110876393', '110517997', '110517890', '110518071', '110837150', '109368630', '110509730', '110509838'], 
                }
recusadas = []
distribuidora = None
competencia = None
ug = None
ug_nova = None
uc_nova = None

def leitor_arquivo():

  global caminho_pdf
  arquivo_aberto = open(caminho_pdf, 'rb')
  arquivo_lido = PyPDF2.PdfReader(arquivo_aberto)
  quantidade_paginas = len(arquivo_lido.pages)

  return arquivo_lido, quantidade_paginas

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
  ug_nova = regex.findall(r'UC.\.?\s*:\s*(\d+)', conteudo)[0]

  competencia = regex.findall(r'Emitido em\s*:\s(\d{2}/\d{2}/\d{4})', conteudo)[0][3:].replace("/", "-")
  for D in distribuidoras.items():
    if encontrar_uc in D[1]:
      distribuidora = D[0]

  tabelas = read_pdf(caminho_pdf, pages=pagina+1, multiple_tables=True, lattice=True, pandas_options={'header': 0})
  tabelas = tabelas[1]
  tabelas.to_csv('{}.csv'.format(encontrar_uc), encoding='utf-8')

  return encontrar_uc, ug_nova

def unidade_consumidora(pagina, arquivo):

  global caminho_pdf
  global uc_nova

  pagina_lida = arquivo.pages[pagina]
  conteudo = pagina_lida.extract_text()
  encontrar_uc = regex.findall(r'UC Conces.\.?\s*:\s*(\d+)', conteudo)[0]
  uc_nova = regex.findall(r'UC.\.?\s*:\s*(\d+)', conteudo)[0]
  
  tabelas = read_pdf(caminho_pdf, pages=pagina+1, multiple_tables=True, lattice=True, pandas_options={'header': 0})
  tabelas = tabelas[1]
  tabelas.to_csv('{}.csv'.format(encontrar_uc), encoding='utf-8')

  return encontrar_uc, uc_nova

def dataframe_pandas(uc, nv, pagina):
  
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
    dados_coluna_novo = [nv, nv, nv, nv, nv, nv, nv, nv, nv, nv, nv, nv, nv, nv]

    df.insert(1, 'Código', dados_coluna_uc)
    df.insert(2, 'Novo Código', dados_coluna_novo)

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
      dados_coluna_novo = [nv, nv, nv, nv, nv, nv, nv, nv, nv, nv, nv, nv, nv, nv]

      df.insert(1, 'Código', dados_coluna_uc)
      df.insert(2, 'Novo Código', dados_coluna_novo)

      df.drop(df.head(1).index, inplace=True)
      df.drop(df.columns[0], axis=1, inplace=True)

      os.remove('{}.csv'.format(uc))

  return df

def save_csv(df):

  global distribuidora
  global competencia
  global ug

  diretorio = 'C:/Users/lucas.conceicao/Relatorio de Compensacao/{}/{}/{}'.format(distribuidora, competencia, ug)
  if not os.path.exists(diretorio):
    os.makedirs(diretorio)
  df.to_excel(os.path.join(diretorio, '{}r.xlsx'.format(ug)), index=False)

def save_csv_alternativo(df):

  global distribuidora
  global competencia
  global ug

  diretorio = 'C:/Users/lucas.conceicao/Relatorio de Compensacao/Diversos/'
  if not os.path.exists(diretorio):
    os.makedirs(diretorio)
  df.to_excel(os.path.join(diretorio, 'Relatório_GD_{}_20251001.xlsx'.format(ug)), index=False)

def relatorio_de_compensacao_organizador():

  dataframe = pandas.DataFrame(columns=['Código', 'Novo Código', 'Referência',
                                        'Saldo Anterior PT','Saldo Anterior FP','Saldo Anterior TP',
                                        'Cred. Receb. Outra UC PT','Cred. Receb. Outra UC FP','Cred. Receb. Outra UC TP',
                                        'Energia Injetada PT','Energia Injetada FP','Energia Injetada TP',
                                        'Energia Ativa PT','Energia Ativa FP','Energia Ativa TP',
                                        'Credito Utilizado no Mês PT','Credito Utilizado no Mês FP','Credito Utilizado no Mês TP',
                                        'Saldo Mes PT','Saldo Mes FP','Saldo Mes TP',
                                        'Saldo Transferido Outra UC PT','Saldo Transferido Outra UC FP','Saldo Transferido Outra UC TP',
                                        'Saldo Final PT','Saldo Final FP','Saldo Final TP'])
  
  arquivo_lido, quantidade_paginas = leitor_arquivo()

  for pagina in range(quantidade_paginas):
    if check_empty(pagina) or check_columns(pagina) == 10:
      pass
    else:
      if pagina == 0:
        codigo_uc, novo = unidade_geradora(pagina, arquivo_lido)
        uc_dataframe = dataframe_pandas(codigo_uc, novo, pagina)
      else:
        codigo_uc, novo = unidade_consumidora(pagina, arquivo_lido)
        uc_dataframe = dataframe_pandas(codigo_uc, novo, pagina)
  
    dataframe = pandas.concat([dataframe, uc_dataframe], axis=0)
    dataframe = dataframe.drop(dataframe.columns[26], axis=1)

    save_csv_alternativo(dataframe)

relatorio_de_compensacao_organizador()