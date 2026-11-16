from tabula import read_pdf
import pandas
import os
import shutil
from openpyxl import load_workbook
from tabulate import tabulate
import time

caminho_txt = 'C:/Users/lucas.conceicao/Downloads/minimicro_000613421603140_20260701.txt'
caminho_recebimento = 'C:/Users/lucas.conceicao/AXS ENERGIA S A/Usinas AXS - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento'
caminho_recusados = "C:/Users/lucas.conceicao/AXS ENERGIA S A/Usinas AXS - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento/Extras/Recusados.xlsx"
caminho_compensados = "C:/Users/lucas.conceicao/AXS ENERGIA S A/Usinas AXS - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento/Extras/Sem_compensacao.xlsx"
distribuidoras = {'COPEL': ['108368980','111254310','111254230','111254167','110069781','110069749','110069692','110022920','110022890','110635710','110519370','110519256','110519329','110635809', '108883663', '108883116', '108883590', '108288897', '108288854', '108288927', '108289010', '000613421603140', '110510526', '110510429', '110510640', '110756517', '110756428', '110756606', '110411889', '110411978', '110411854', '110877306', '110877519', '110877446', '110877845', '110877918', '110878019', '110876806', '110876865', '110876393', '110517997', '110517890', '110518071', '110837150', '109368630', '110509730', '110509838'], }

dataframe_maior = None
maior = []
dataframe_meio = None
meio = []
dataframe_menor = None
menor = []
erro = []
dataframe_final = None

recusadas = []
energia_recusadas = []

distribuidora = 'Copel'
competencia = '07_2026'
nome_arquivo = caminho_txt.split('/')[-1]
ug = str(nome_arquivo.split('_')[1])

geracao_ug = None
consumo_ug = None
transferencia_ug = None

geracao_calc = None
consumo_calc = None
transferido_calc = None

dataframe_aprovados = None
sem_compensacao = None
sem_recebimento = None
dataframe_conferidos = None
dataframe_recusados = None
dataframe_divergencias = None


def dataframe_pandas_copel():

  global dataframe_aprovados
  global sem_compensacao
  global sem_recebimento
  global dataframe_final
  global dataframe_maior
  global dataframe_menor
  global maior
  global menor
  global erro
  global geracao_ug
  global consumo_ug
  global transferencia_ug

  cabecalho_10 = None
  cabecalho_23 = None
  cabecalho_30 = None

  with open(caminho_txt, 'r', encoding='latin-1') as f:

    for linha in f:
      
      linha_lista = linha.strip().split(';')
      tamanho = len(linha_lista)
      #print(tamanho)
      is_header = "UC Ger." in linha_lista

      if tamanho == 10:
          
          if is_header:
            cabecalho_10 = linha_lista
          else:
             menor.append(linha_lista)
      elif tamanho == 23:
          if is_header:
             cabecalho_23 = linha_lista
          else:
            meio.append(linha_lista)
      elif tamanho == 30:
          if is_header:
             cabecalho_30 = linha_lista
          else:
            maior.append(linha_lista)
      else:
          erro.append(linha_lista)

  dataframe_menor = pandas.DataFrame(menor, columns=cabecalho_10)
  dataframe_meio = pandas.DataFrame(meio, columns=cabecalho_23)
  dataframe_maior = pandas.DataFrame(maior, columns=cabecalho_30)

  dataframe_final = pandas.concat([dataframe_meio, dataframe_menor], ignore_index=True)
  dataframe_final = pandas.concat([dataframe_maior, dataframe_final], ignore_index=True)
  
  dataframe_final = dataframe_final.fillna(0)
  dataframe_final = dataframe_final[dataframe_final['Mês Ref.']=='01/07/2026']

  geracao_ug = int(dataframe_final[dataframe_final['UC Ger.']==ug]['Energia Injetada Ponta'].values[0]) + int(dataframe_final[dataframe_final['UC Ger.']==ug]['Energia Injetada Fora Ponta'].values[0])
  
  consumo_ug = int(dataframe_final[dataframe_final['UC Ger.']==ug]['Energia Ativa Ponta'].values[0]) + int(dataframe_final[dataframe_final['UC Ger.']==ug]['Energia Ativa Fora Ponta'].values[0])
  transferido_ug = int(dataframe_final[dataframe_final['UC Ger.']==ug]['Saldo Transferido Outra UC Ponta'].values[0]) + int(dataframe_final[dataframe_final['UC Ger.']==ug]['Saldo Transferido Outra UC Fora Ponta'].values[0])

  dataframe_aprovados = dataframe_final[(dataframe_final['Crédito recebido outra UC Fora Ponta'] != "0")]

  sem_compensacao = dataframe_final[(dataframe_final['Energia Ativa Todos os Períodos'] != "0") & (dataframe_final['Crédito recebido outra UC Fora Ponta'] == "0") & (dataframe_final['Crédito Utilizado no Mês Todos os Períodos'] == "0")]
  sem_compensacao = sem_compensacao[['Mês Ref.', 'UC']]

  sem_recebimento = dataframe_final[(dataframe_final['Crédito recebido outra UC Fora Ponta'] == "0")]
  sem_recebimento = sem_recebimento[['Mês Ref.', 'UC']]

  dataframe_final.to_excel(f'{ug}.xlsx')
  #print(dataframe_aprovados)

def dataframe_pandas_recebimento(data):
  
  global caminho_recebimento
  global ug
  global geracao_calc
  global consumo_calc
  global transferido_calc
  
  arquivo = "{} - COPEL.xlsm".format(data)
  caminho_recebimento_arq = os.path.join(caminho_recebimento, arquivo)
  dataframe = pandas.read_excel(caminho_recebimento_arq, ug)

  consumo_calc = int(dataframe[dataframe['Instalacao'] == int(ug)]['Qtd_consumo'].sum())
  geracao_calc = int(dataframe[dataframe['Instalacao'] == int(ug)]['Qtd_geracao'].sum())
  transferido_calc = int(dataframe[dataframe['Instalacao'] == int(ug)]['Qtd_transferencia'].sum())

  dataframe = dataframe.drop(index=0)
  dataframe = dataframe.drop(index=len(dataframe))

  return dataframe

def dataframe_pandas_comparativo(df_recebimento):
  
  global dataframe_conferidos
  global recusadas
  global energia_recusadas
  global dataframe_recusados

  dataframe_conferidos = pandas.DataFrame(columns=['Código', 'Recebimento AXS',
                                        'Recebimento Distribuidora', 'Diferenca', 'variacao'])

  for index, row in df_recebimento.iterrows():
      if str(row['Instalacao']) in dataframe_aprovados['UC'].values:
          codigo = int(row['Instalacao'])
          Recebimento_AXS = row['Qtd_recebimento']
          Recebimento_Distribuidora = int(dataframe_aprovados[dataframe_aprovados['UC'] == str(row['Instalacao'])]['Crédito recebido outra UC Ponta'].values[0]) + int(dataframe_aprovados[dataframe_aprovados['UC'] == str(row['Instalacao'])]['Crédito recebido outra UC Fora Ponta'].values[0])
          Diferenca = int(Recebimento_AXS)-int(Recebimento_Distribuidora)
          variacao = Diferenca/int(Recebimento_Distribuidora)*100
          dataframe_conferidos.loc[len(dataframe_conferidos)] = [codigo, Recebimento_AXS, Recebimento_Distribuidora, Diferenca, variacao]
      else:
          recusadas.append(row['Instalacao'])
          energia_recusadas.append(row['Qtd_recebimento'])
  
  dataframe_recusados = pandas.DataFrame({'Comp':competencia, 'Distribuidora':'Copel', 'Usina':'Usina', 'UG':ug, 'UC': recusadas, 'Motivo': 'Sem Informação', 'Energia': energia_recusadas})

def validacao_medicoes():

  if consumo_calc == consumo_ug and geracao_calc == geracao_ug:

    print('                  Medições Digitadas Corretas!!! \n')
    print('Consumo Digitado: {}'.format(consumo_calc))
    print('Consumo Real: {} \n'.format(consumo_ug))

    print('Geração Digitado: {}'.format(geracao_calc))
    print('Geração Real: {} \n'.format(geracao_ug))

    return True
  
  else:
  
    print('                  Medições Digitadas Incorretas!!! \n')
    print('Consumo Digitado: {}'.format(consumo_calc))
    print('Consumo Real: {}'.format(consumo_ug))

    print('Geração Digitado: {}'.format(geracao_calc))
    print('Geração Real: {}'.format(geracao_ug))

    return False

def validacao_recebimentos(df_comparativo):

  global dataframe_divergencias

  dataframe_divergencias = df_comparativo[(df_comparativo['Diferenca'] > 3) | (df_comparativo['Diferenca'] < -3)]
  divergencias = dataframe_divergencias.shape[0]

  if divergencias == 0:
    return True
  else:
    False

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

  origem = caminho_txt
  destino = os.path.join(diretorio, caminho_txt.split('/')[-1])
  shutil.move(origem, destino)

def recusados ():

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

def nao_compensadoss ():

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

def relatorio_de_compensacao_organizador(data):
  
  global geracao_ug
  global consumo_ug
  global transferencia_ug
  global dataframe_recusados
  global sem_compensacao

  dataframe_pandas_copel()
  
  dataframe_recebimento = dataframe_pandas_recebimento(data)
  dataframe_recebimento['Instalacao'] = dataframe_recebimento['Instalacao'].astype(int)

  if validacao_medicoes():

    dataframe_pandas_comparativo(dataframe_recebimento)
  
    if validacao_recebimentos(dataframe_conferidos):
      print("                     Todos os créditos recebidos pelos beneficiários estão corretos \n ")
      print('                                        RESUMO DOS BENEFICIÁRIOS \n')
      print('                                ************  APROVADOS  ************ \n')
      colunas = dataframe_conferidos.columns.tolist()
      print(tabulate(dataframe_conferidos, headers=colunas, tablefmt='fancy_grid'))
      print(' ')
      print('                                 ***********  RECUSADOS  ************')
      colunas = dataframe_recusados.columns.tolist()
      print(tabulate(dataframe_recusados, headers=colunas, tablefmt='fancy_grid'))
      print(' ')
      print('                              ***********  SEM COMPENSAÇÃO  ************\n')
      colunas = sem_compensacao.columns.tolist()
      print(tabulate(sem_compensacao, headers=colunas, tablefmt='fancy_grid'))
      print(' ')

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


relatorio_de_compensacao_organizador('JUL26')

