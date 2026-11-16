import pandas
import os
import datetime
import shutil
from openpyxl import load_workbook
from tabulate import tabulate
from dateutil.relativedelta import relativedelta

# Variáveis Globais

# Caminhos
arquivo = 'Relatório_GD_3014900405_20261112'
caminho_arquivo = 'C:/Users/lucas.conceicao/Downloads'
caminho_depara = 'C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/03.Análises/DEPARA_CEMIG.xlsx'
caminho_analise_de_faturas = 'C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Faturas/Análise de Faturas.xlsm'
caminho_lista_vigente = 'C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/Arquivos BI/Controle de Estoque - USINAS/Base previsão faturamento - Lista Vigente.xlsm'
caminho_recebimentos = 'C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento'
caminho_conferencias = 'C:/Users/lucas.conceicao/Relatorio de Compensacao'
caminho_arquivo_c = None
caminho_recusados = "C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento/Extras/Recusados BKP.xlsx"
caminho_compensados = "C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento/Extras/Sem_compensacao.xlsx"
# Estruturas de Consulta

# Estruturas de Armazenamento
lista_recusados_uc = []
lista_recusados_energia = []

# Variáveis de Armazenamento
dicio_distribuidora = {'CEMIG': ['3015948206', '3014965842', '3014965872', '3014965846', '3014306059', '3014868693','3014868655','3014784111','3014784110','3014784085','3015056747','3014583262','3014583260','3014583209','3014784110','3014784085','3014583209', '3014629309', '3014629582', '3014629614', '3014629370', '3014109500', '3014109501', '3014109502', '3014233708', '3015083428', '3015083492', '3014926608', '3014926612', '3014233705', '3014298175', '3014298176', '3014427046', '3015120874', '3014292343', '3014292356', '3014292382', '3014837599', '3014837605', '3014837606', '3014902567', '3014902579', '3014902584', '3014900389', '3014900403', '3014900405', '3014868833', '3014868836','3014868860']}
dicio_meses = {'1': 'JAN', '2': 'FEV', '3': 'MAR', '4': 'ABR', '5': 'MAI', '6': 'JUN', '7': 'JUL', '8': 'AGO', '9': 'SET', '10': 'OUT', '11': 'NOV', '12': 'DEZ'}
dicio_ano = {'2024': '24', '2025': '25', '2026': '26'}

distribuidora = None
competencia = None
ug = None
ug_nova = None

consumo_ug = None
geracao_ug = None
transferencia_ug = None

consumo_calc = None
geracao_calc = None
compensado_calc = None
transferido_calc = None

dataframe_axs = None
dataframe_exportado = None
dataframe_deapara = pandas.read_excel(caminho_analise_de_faturas, sheet_name="NOVOS No. UG")
dataframe_deapara = dataframe_deapara[["Antiga UG", "Nova UG"]]
de_para_dict = dict(zip(dataframe_deapara["Antiga UG"],dataframe_deapara['Nova UG']))
dataframe_aprovados = None
dataframe_reprovados = None
dataframe_nao_previstos = None
dataframe_nao_compensados = None
dataframe_divergencias = None
dataframe_analise = None

total_rec = 0
total_dis = 0
contador_aprovados = 0
contador_reprovados = 0
contador_nao_previstos = 0
contador_nao_compensados = 0
sem_compensacao = None

def conferir_arquivo (comp):

    global arquivo
    global caminho_conferencias
    global distribuidora
    global competencia
    global ug
    global caminho_arquivo_c

    competencia = comp.strftime('%m-%Y')
    ug = arquivo[13:23]
    distribuidora = [k for k,v in dicio_distribuidora.items() if ug in v][0]
    
    arquivo_c = '{}c.xlsx'.format(ug)
    pasta = '{}/{}/{}/{}'.format(caminho_conferencias, distribuidora, competencia, ug)
    caminho_arquivo_c = os.path.join(pasta, arquivo_c)

    if os.path.exists(caminho_arquivo_c):
        print('A conferência da UG {} já foi realizada'.format(os.path.basename(caminho_arquivo_c)))
        return True
    else:
        return False

def dataframe_distribuidora (comp): 
    
    global consumo_ug
    global geracao_ug
    global transferencia_ug

    arquivo_d = '{}.xlsx'.format(arquivo)
    caminho_arquivo_d = os.path.join(caminho_arquivo, arquivo_d)

    dataframe = pandas.read_excel(caminho_arquivo_d)

    competencia = '{}/{}'.format(comp.year, comp.month-1)

    dataframe = dataframe[dataframe['Periodo'] == competencia]
    dataframe = dataframe.sort_values("Qtd_compensacao").drop_duplicates(subset = ['Periodo', 'Periodo_uso', 'Modalidade', 'Instalacao', 'Quota', 'Posto_horario'] ,keep="last").reset_index()

    consumo_ug = dataframe[dataframe['Instalacao'] == int(ug)]['Qtd_consumo'].sum()
    geracao_ug = dataframe[dataframe['Instalacao'] == int(ug)]['Qtd_geracao'].sum()
    transferencia_ug = dataframe[dataframe['Instalacao'] == int(ug)]['Qtd_transferencia'].sum()

    dataframe = dataframe[dataframe['Instalacao'] != int(ug)]
    dataframe = dataframe.sort_values("Instalacao").reset_index()

    dataframe = dataframe.drop(columns=['level_0', 'index', 'Periodo', 'Modalidade', 'Posto_horario', 'Qtd_geracao', 'Qtd_transferencia', 'Qtd_saldo_ant', 'Qtd_saldo_atual', 'Qtd_saldo_exp', 'Qtd_prox_saldo_exp', 'Per_prox_saldo_exp'])

    dataframe['Numero_uc'] = dataframe['Numero_uc'].astype(str).str.replace(r'[^\d]', '', regex=True)

    dataframe['Numero_uc'] = dataframe['Numero_uc'].astype(int)

    return dataframe

def dataframe_analise_de_faturas (data):

    global dataframe_analise

    dataframe = pandas.read_excel(caminho_analise_de_faturas, sheet_name="Geral", skiprows=2)

    data_ajusta = f'{str(data)[0:5]}-{int(str(data)[5:7])-1}-01'

    dataframe = dataframe[(dataframe["Mês"] == data_ajusta) & (dataframe["Distribuidora"] == "Cemig")]

    dataframe = dataframe[["Mês", "UG", "Cons P Reg", "Cons FP Reg", "Inj P Reg", "Inj FP Reg", "Saldo Anterior P kWh", "Saldo Anterior FP kWh",]]

    dataframe_analise = dataframe
    
    return dataframe_analise

def excedente_geracao ():

    global ug_nova
    global dataframe_tarifas
    global geracao_calc
    global consumo_calc
    global compensado_calc
    global transferido_calc

    ug_nova = de_para_dict[int(ug)]

    dataframe = dataframe_analise[dataframe_analise["UG"]==ug_nova]
    
    consumo_ponta = dataframe["Cons P Reg"].values[0]
    consumo_fora_ponta = dataframe["Cons FP Reg"].values[0]

    injetado_ponta = dataframe["Inj P Reg"].values[0]
    injetado_fora_ponta = dataframe["Inj FP Reg"].values[0]

    consumo_calc = round(consumo_ponta + consumo_fora_ponta)
    geracao_calc = round(injetado_ponta + injetado_fora_ponta)

    credito_utilizado_ponta = (consumo_ponta)
    credito_utilizado_fora_ponta = (consumo_fora_ponta)

    compensado_calc = round(credito_utilizado_ponta + credito_utilizado_fora_ponta)
    excedente = round((injetado_ponta + injetado_fora_ponta) - (credito_utilizado_ponta + credito_utilizado_fora_ponta))

    transferido_calc = round(excedente)

    return excedente

def dataframe_recebimentos (comp, exc):
    
    data_inicial = f'{str(comp)[0:5]}-{int(str(comp)[5:7])-1}-01'
    
    dataframe_axs = pandas.read_excel(caminho_lista_vigente, "LISTA VIGÊNTE")
    
    dataframe_axs = dataframe_axs[(dataframe_axs["UG"] == int(ug)) & (dataframe_axs["COMPETÊNCIA"] == data_inicial)]
    
    dataframe_axs["Qtd_recebimento"] = dataframe_axs["Quota"] * exc

    dataframe_axs = dataframe_axs[["COMPETÊNCIA", "UG", "UC", "Quota", "Qtd_recebimento"]]

    return dataframe_axs

def validacao_medicoes_ug ():

    if consumo_calc == consumo_ug and geracao_calc == geracao_ug and transferido_calc == transferencia_ug:
        print('Medições Digitadas Corretas!!! \n')
        print('Consumo Digitado: {}'.format(consumo_calc))
        print('Consumo Real: {}'.format(consumo_ug))

        print('Geração Digitado: {}'.format(geracao_calc))
        print('Geração Real: {}'.format(geracao_ug))

        print('Transferência Digitada: {}'.format(transferido_calc))
        print('Transferência Real: {}'.format(transferencia_ug))

        return True
    else:
        print('Medições Digitadas Incorretas!!!')
        print('Consumo Digitado: {}'.format(consumo_calc))
        print('Consumo Real: {}'.format(consumo_ug))

        print('Geração Digitado: {}'.format(geracao_calc))
        print('Geração Real: {}'.format(geracao_ug))

        print('Transferência Digitada: {}'.format(transferido_calc))
        print('Transferência Real: {}'.format(transferencia_ug))

        return False

def validacao_recebimentos (df_recebimentos, df_distribuidora):

    global dataframe_aprovados
    global dataframe_reprovados
    global dataframe_nao_previsto
    global dataframe_nao_compensados
    global dataframe_divergencias
    global total_rec
    global total_dis
    global contador_aprovados
    global contador_reprovados
    global contador_nao_previstos
    global contador_nao_compensados
    
    dataframe_aprovados = pandas.DataFrame(columns=['Data', 'Código', 'Quota', 'Recebimento AXS', 'Recebimento Distribuidora', 'Diferenca', 'variacao'])
    
    dataframe_reprovados = pandas.DataFrame(columns=['Data', 'Código', 'Recebimento AXS'])
    
    dataframe_nao_previsto = pandas.DataFrame(columns=['Data', 'Código', 'Recebimento Distribuidora'])

    dataframe_nao_compensados = pandas.DataFrame(columns=['Data', 'Código', 'Recebimento Distribuidora'])

    contador_aprovados = 0
    contador_reprovados = 0

    for index, row in df_recebimentos.iterrows():
        data = row['COMPETÊNCIA']
        instalacao = row['UC']
        cota = row['Quota']
        recebimento_axs = round(row['Qtd_recebimento'], 5)
    
        if instalacao in df_distribuidora['Instalacao'].values:
            
            recebimento_distribuidora = round(df_distribuidora[df_distribuidora['Instalacao'] == instalacao]['Qtd_recebimento'].sum(), 5)
            diferenca = round(recebimento_axs - recebimento_distribuidora, 5)

            if diferenca == 0:
                dataframe_aprovados.loc[len(dataframe_aprovados)] = [data, instalacao, cota, recebimento_axs, recebimento_distribuidora, diferenca, f'0%']
            else:
                variacao = round((diferenca/recebimento_axs) * 100, 4)
                dataframe_aprovados.loc[len(dataframe_aprovados)] = [data, instalacao, cota, recebimento_axs, recebimento_distribuidora, diferenca, f'{variacao}%']
            contador_aprovados += 1
        
        elif instalacao in df_distribuidora['Numero_uc'].values:

            recebimento_distribuidora = round(df_distribuidora[df_distribuidora['Numero_uc'] == instalacao]['Qtd_recebimento'].sum(), 5)
            diferenca = round(recebimento_axs - recebimento_distribuidora, 5)

            if diferenca == 0:
                dataframe_aprovados.loc[len(dataframe_aprovados)] = [data, instalacao, cota, recebimento_axs, recebimento_distribuidora, diferenca, f'0%']
            else:
                variacao = round((diferenca/recebimento_axs) * 100, 4)
                dataframe_aprovados.loc[len(dataframe_aprovados)] = [data, instalacao, cota, recebimento_axs, recebimento_distribuidora, diferenca, f'{variacao}%']
            contador_aprovados += 1

        else:
            dataframe_reprovados.loc[len(dataframe_reprovados)] = [data, instalacao, recebimento_axs]
            contador_reprovados += 1

    instalacoes_dist = set(df_distribuidora['Instalacao'])
    instalacoes_rece = set(df_recebimentos['UC'])
    
    dataframe_nao_compensados = df_distribuidora.groupby('Instalacao').sum()
    dataframe_nao_compensados = dataframe_nao_compensados[(dataframe_nao_compensados['Qtd_recebimento'] != 0) & (dataframe_nao_compensados['Qtd_consumo'] != 0) & (dataframe_nao_compensados['Qtd_compensacao'] == 0)]
    contador_nao_compensados = dataframe_nao_compensados.shape[0]

    total_rec = len(instalacoes_rece)
    total_dis = len(instalacoes_dist)

    instalacoes_nao_previstas = instalacoes_dist.difference(instalacoes_rece)
    
    for index in instalacoes_nao_previstas:

        linha_instalacao = df_distribuidora[df_distribuidora['Instalacao'] == index]
        data = linha_instalacao['Periodo_uso']
        recebimento_distribuidora = linha_instalacao['Qtd_recebimento'].sum()
        dataframe_nao_previsto.loc[len(dataframe_nao_previsto)] = [data, index, recebimento_distribuidora]
        
    dataframe_divergencias = dataframe_aprovados[(dataframe_aprovados['Diferenca'] >= 1) | (dataframe_aprovados['Diferenca'] <= -1)]
    divergencias = dataframe_divergencias.shape[0]

    if divergencias == 0:
        return True
    
    else:
        return False

def ajustes_de_credito (df_recebimentos):

    for index, row in dataframe_divergencias.iterrows():

        codigo = row["Código"]
        credito = row["Recebimento Distribuidora"]

        df_recebimentos.loc[df_recebimentos["UC"] == codigo, "Qtd_recebimento"] = credito

def salvar (df_a, df_r, df_n, df_nc):

    global caminho_arquivo_c

    caminho_c = 'C:/Users/lucas.conceicao/OneDrive 2026/AXS ENERGIA S A/Gestão de Usinas - Usinas AXS/03.Análises/14 - Controle de Faturas e Recebimentos/Recebimento/Relatórios de Recebimentos/{}/{}/{}'.format(distribuidora, competencia, ug)
    os.makedirs(caminho_c)

    caminho_arquivo_c = os.path.join(caminho_c, '{}c.xlsx'.format(ug))
    
    with pandas.ExcelWriter(caminho_arquivo_c, engine='xlsxwriter') as writer:

        df_a.to_excel(writer, sheet_name='Aprovados', index=False)
        df_r.to_excel(writer, sheet_name='Reprovados', index=False)
        df_n.to_excel(writer, sheet_name='Nao Previstos', index=False)
        df_nc.to_excel(writer, sheet_name='Não Compensados', index=False)

    print(f"Arquivo {os.path.basename(caminho_arquivo_c)} criado com sucesso!")

    origem = os.path.join(caminho_arquivo, '{}.xlsx'.format(arquivo))
    destino = os.path.join(caminho_c, '{}.xlsx'.format(arquivo))
    shutil.move(origem, destino)
    
    print("Arquivo {}.xlsx movido com sucesso!".format(arquivo))

def csv (data: datetime, ug: str):

    global dataframe_exportado

    data_hoje = data.strftime('%Y.%m.%d')

    dataframe = pandas.DataFrame(columns=['Periodo', 'Modalidade', 'Instalacao', 'Quota', 'Posto_horario', 'Qtd_consumo', 'Qtd_geracao', 'Qtd_compensacao', 'Qtd_saldo_ant', 'Qtd_transferencia', 'Qtd_recebimento', 'Qtd_saldo_atual', 'Qtd_saldo_exp', 'Qtd_prox_saldo_exp', 'Per_prox_saldo_exp'])

    Periodo = str(data)[:7].replace("-", "/")
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
    
    dataframe.loc[len(dataframe)] = [Periodo, Modalidade, Instalacao, Quota, Posto_horario, Qtd_consumo, Qtd_geracao, Qtd_compensacao, Qtd_saldo_ant, Qtd_transferencia, Qtd_recebimento, Qtd_saldo_atual, Qtd_saldo_exp, Qtd_prox_saldo_exp, Per_prox_saldo_exp]
    
    for index, row in dataframe_aprovados.iterrows():
        Periodo = str(data)[:7].replace("-", "/")
        Modalidade = "Ger. Compartilhada-Recebedora"
        Instalacao = str(round(row["Código"]))
        Quota = row["Quota"]
        Posto_horario = "FORAP"
        Qtd_consumo = ""
        Qtd_geracao = ""
        Qtd_compensacao = ""
        Qtd_saldo_ant = ""
        Qtd_transferencia = ""
        Qtd_recebimento = str(row["Recebimento AXS"]).replace(".", ",")
        Qtd_saldo_atual = ""
        Qtd_saldo_exp = ""
        Qtd_prox_saldo_exp = ""
        Per_prox_saldo_exp = ""

        dataframe.loc[len(dataframe)] = [Periodo, Modalidade, Instalacao, Quota, Posto_horario, Qtd_consumo, Qtd_geracao, Qtd_compensacao, Qtd_saldo_ant, Qtd_transferencia, Qtd_recebimento, Qtd_saldo_atual, Qtd_saldo_exp, Qtd_prox_saldo_exp, Per_prox_saldo_exp]

    Periodo = str(data)[:7].replace("-", "/")
    Modalidade = "Ger. Compartilhada-Geradora"
    Instalacao = ug
    Quota = 0
    Posto_horario = "FORAP"
    Qtd_consumo = consumo_calc
    Qtd_geracao = geracao_calc
    Qtd_compensacao = compensado_calc
    Qtd_saldo_ant = ""
    Qtd_transferencia = transferido_calc
    Qtd_recebimento = ""
    Qtd_saldo_atual = ""
    Qtd_saldo_exp = ""
    Qtd_prox_saldo_exp = ""
    Per_prox_saldo_exp = ""
    
    dataframe.loc[len(dataframe)] = [Periodo, Modalidade, Instalacao, Quota, Posto_horario, Qtd_consumo, Qtd_geracao, Qtd_compensacao, Qtd_saldo_ant, Qtd_transferencia, Qtd_recebimento, Qtd_saldo_atual, Qtd_saldo_exp, Qtd_prox_saldo_exp, Per_prox_saldo_exp]
        
    dataframe["Quota"] = dataframe["Quota"].astype(float) / 100

    nome_arquivo = '{} {}.csv'.format(str(data_hoje), ug)

    caminho_arquivo = os.path.join('C:/Users/lucas.conceicao/Downloads', nome_arquivo)

    dataframe.to_csv(caminho_arquivo, index=False, sep=";", encoding="utf-8-sig", decimal=',')

    dataframe_exportado = dataframe

    print(f"Arquivo CSV {os.path.basename(caminho_arquivo)} criado com sucesso!")

def recusados ():

    pergunta = input("Deseja realizar a função de adicionar os recusados no arquivo de conferência na pasta de destino: ")
    
    if pergunta.lower() == "s":

        recusado = []
        try:
            workbook = load_workbook(caminho_recusados)
            sheet = workbook['Historico']

            for row in dataframe_reprovados.itertuples():
                recusado.append(row[1])
                recusado.append(distribuidora)
                recusado.append('UFV_NOME')
                recusado.append(ug)
                recusado.append(row[2])
                recusado.append('Sem informação')
                recusado.append(row[3])

                sheet.append(recusado)
                workbook.save(caminho_recusados)

                recusado = []

                print("Clientes recusados foram adicionados no arquivo de referência")

        except FileNotFoundError:
            print(f"O arquivo não foi encontrato.")
        except Exception as e:
            print(f"Ocorreu um erro: {e}")
    else:
            print("Os recusados não foram adicionados na base de 'recusados'")

def sem_compesacao ():

    pergunta = input("Deseja realizar a função de adicionar os sem compensação no arquivo de conferência na pasta de destino: ")
    
    if pergunta.lower() == "s":
        
        nao_compensados = []
        try:
            workbook = load_workbook(caminho_compensados)
            sheet = workbook['Planilha1']

            for row in dataframe_nao_compensados.itertuples():
                nao_compensados.append(row[1])
                nao_compensados.append(distribuidora)
                nao_compensados.append('UFV_NOME')
                nao_compensados.append(ug)
                nao_compensados.append(row[2])
                nao_compensados.append('Sem informação')
                nao_compensados.append(row[3])

                sheet.append(nao_compensados)
                workbook.save(caminho_compensados)

                nao_compensados = []

                print("Clientes sem compensação foram adicionados no arquivo de referência")

        except FileNotFoundError:
            print(f"O arquivo não foi encontrato.")
        except Exception as e:
            print(f"Ocorreu um erro: {e}")

    else:
        print("Os sem compensação não foram adicionados na base de 'Sem Compensação'")

def main (comp=datetime.date.today()):

    global dataframe_axs

    if not conferir_arquivo(comp):
        df_distribuidora = dataframe_distribuidora(comp)
        dataframe_analise_de_faturas(comp)
        excendente = excedente_geracao()
        dataframe_axs = dataframe_recebimentos(comp, excendente)
        
        if validacao_medicoes_ug():

            def secundario():

                if validacao_recebimentos(dataframe_axs, df_distribuidora):

                    print("                     Todos os créditos recebidos pelos beneficiários estão corretos \n ")
                    print('                                    RESUMO DOS BENEFICIÁRIOS \n')
                    print('     Planilha de Recebimento  /  Demonstrativo de Compensação')
                    print('  TOTAL                  {}      {}'.format(total_rec, total_dis))
                    print('APROVADOS                {}      {}'.format(contador_aprovados, contador_aprovados))
                    print('REPROVADOS                {}'.format(contador_reprovados))
                    print('INDEVIDOS                          {}'.format(contador_nao_previstos))
                    print('NÃO COMPENSADOS           {} \n '.format(contador_nao_compensados))
                    print('                          ************  APROVADOS  ************')
                    colunas = dataframe_aprovados.columns.tolist()
                    print(tabulate(dataframe_aprovados, headers=colunas, tablefmt='fancy_grid'))
                    print(' ')
                    print('          ***********  REPROVADOS  ************')
                    colunas = dataframe_reprovados.columns.tolist()
                    print(tabulate(dataframe_reprovados, headers=colunas, tablefmt='fancy_grid'))
                    #print('\n                          **********  NÃO PREVISTOS  **********')
                    #colunas = dataframe_nao_previsto.columns.tolist()
                    #print(tabulate(dataframe_nao_previsto, headers=colunas, tablefmt='fancy_grid'))
                    print(' \n                         **********  NÃO COMPENSADOS  **********')
                    colunas = dataframe_nao_compensados.columns.tolist()
                    print(tabulate(dataframe_nao_compensados, headers=colunas, tablefmt='fancy_grid'))

                    pergunta = input("Deseja realizar a função de criação e salvamento do arquivo de conferência na pasta de destino: ")

                    if pergunta.lower() == "s":
                        salvar(dataframe_aprovados, dataframe_reprovados, dataframe_nao_previsto, dataframe_nao_compensados)
                        csv(comp, ug)
                        print(f'dataframe aprovados: {len(dataframe_aprovados)}')
                        print(f'dataframe exportado: {len(dataframe_exportado)-2} \n')
                    else:
                        print("Os arquivo não foi criado")
                    
                    if len(dataframe_reprovados) != 0:
                        recusados()

                    if len(dataframe_nao_compensados) != 0:
                        sem_compesacao()

                else:
                    print('                               Há divergências de alguns créditos \n \n ')
                    print('                                    RESUMO DOS BENEFICIÁRIOS \n \n ')
                    colunas = dataframe_divergencias.columns.tolist()
                    print(tabulate(dataframe_divergencias, headers=colunas, tablefmt='fancy_grid'))
                    print("Não foi criado o arquivo de conferência")

                    pergunta = input("Deseja realizar a função de ajuste dos créditos: ")
                                            
                    if pergunta.lower() == "s":
                        ajustes_de_credito(dataframe_axs)
                        secundario()
                    else:
                        print("Não foi criado o arquivo de conferência")

            secundario()


main()