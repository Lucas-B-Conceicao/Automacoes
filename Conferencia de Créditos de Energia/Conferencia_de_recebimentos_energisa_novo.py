import os
import re
import shutil
import tempfile
from tabulate import tabulate
from datetime import datetime
import pandas as pd
import pdfplumber
import calendar

MAPA_NOVA_PARA_ANTIGA = {
    "1.941.939.017-62": "4030758",
    "1.941.938.017-77": "4030767",
    "1.941.940.017-25": "4030749",
    "1.941.950.017-07": "4030691",
    "1.941.958.017-30": "4030658",
    "1.941.968.017-12": "4030594",
    "1.962.316.017-60": "4200568",
    "871.084.017-42": "4200527",
    "1.962.349.017-81": "4200398",
    "955.455.017-79": "4385076",
    "275.546.017-41": "4385297",
    "2.070.805.017-03": "4761002",
    "2.054.945.017-55": "4756717",
    "1.982.149.017-47": "4509670",
    "1.416.330.017-38": "4761010",
    "1.314.586.017-60": "4509657",
    "1.192.318.017-45": "4509662",
    "1.455.682.017-04": "4401020",
    "1.237.069.017-60": "4401014",
    "2.000.159.017-90": "4400952",
    "815.409.017-60": "4563113",
    "382.776.017-48": "4563099",
    "382.777.017-33": "4563050",
    "667.973.017-60": "4868282",
    "882.913.017-53": "4868289",
    "1.404.417.017-03": "4868296",
    "18.890.017-09": "4556228",
    "1.224.855.017-57": "4556218",
    "2.016.632.017-98": "4556154",
    "362.232.017-86": "4243410",
    "2.001.766.017-02": "4243527",
    "2.090.419.017-76": "4909902",
    "2.062.197.017-91": "4926433",
    "1.102.319.017-42": "4518372",
    "1.983.148.017-16": "4385377",
    "2.062.397.017-38": "4909992",
    "1.156.316.017-01": "4518402",
    "979.296.017-05": "4518482",
    "2.001.770.017-39": "4243513",
}

PASTA_USUARIO = os.path.expanduser('~')

PDF_CAMINHO = os.path.join(
    PASTA_USUARIO, 
    "Downloads",
    "demonstrativo_compensacao.pdf"
)

DIRETORIO_RAIZ = os.path.join(
    PASTA_USUARIO,
    "OneDrive 2026",
    "AXS ENERGIA S A", 
    "Gestão de Usinas - Usinas AXS"
                            )

EXCEL_DIR = os.path.join(
            DIRETORIO_RAIZ,
            "03.Análises", 
            "14 - Controle de Faturas e Recebimentos", 
            "Faturas"
                        )

EXCEL_2_DIR = os.path.join(
            DIRETORIO_RAIZ,
            "Arquivos BI",
            "Controle de Estoque - USINAS"
)

DEPARA_DIR = os.path.join(
            DIRETORIO_RAIZ,
            "03.Análises"
                        )

AUDITORIA_DIR = os.path.join(
            DEPARA_DIR,
            "14 - Controle de Faturas e Recebimentos",
            "Recebimento",
            "Relatórios de Recebimentos",
            "ENERGISA"
                        )


dicio_meses = {'JAN': '1', 'FEV': '2', 'MAR': '3', 'ABR': '4', 'MAI': '5', 'JUN': '6', 'JUL': '7', 'AGO': '8', 'SET': '9', 'OUT': '10', 'NOV': '11', 'DEZ': '12'}
dicio_ano = {'24': '2024', '25': '2025', '26': '2026'}

dataframe_axs = None
dataframe_energisa = None

dataframe_aprovados = None
dataframe_reprovados = None
dataframe_divergencias = None
dataframe_nao_esperados = None
dataframe_analise = None

clientes_esperados = 0
clientes_utilizados = 0

contador_aprovados = 0
contador_reprovados = 0

energia_recusada = None

medicoes_axs = {"consumo_axs" : None, "geracao_axs" : None, "compensado_axs": None,"transferido_axs" : None}
medicoes_distribuidora = None

# ---------------------------------------------------------------------
# 1. PREPARAÇÃO DO DATAFRAME DO RELATÓRIO DE COMPENSAÇÃO DA DISTRIBUIDORA
# ---------------------------------------------------------------------

def medicao_distribuidora(caminho_pdf: str) -> dict:

    dict_medicoes = {
        "Mês Referência": None,
        "Consumo Pt (kWh)": 0,
        "Consumo FPt (kWh)": 0,
        "Consumo Total (kWh)": 0,
        "Injetado Pt (kWh)": 0,
        "Injetado FPt (kWh)": 0,
        "Injetado Total (kWh)": 0,
        "Compensado Pt (kWh)": 0,
        "Compensado FPt (kWh)": 0,
        "Compensado Total (kWh)": 0,
    }

    def limpar_numero(val: str) -> int:
        try:
            val_clean = val.replace(".", "").replace(",", ".")
            return int(float(val_clean))
        except (ValueError, TypeError):
            return 0

    with pdfplumber.open(caminho_pdf) as pdf:
        texto_pagina1 = pdf.pages[0].extract_text()

        # 1. Mês de Referência (ex: "08/2026")
        match_ref = re.search(r"Referência:\s*(\d{2}/\d{4})", texto_pagina1)
        if match_ref:
            dict_medicoes["Mês Referência"] = match_ref.group(1)

        linhas = texto_pagina1.split("\n")

        em_consumo_ponta = False
        em_consumo_fora_ponta = False
        posto_atual = None

        for linha in linhas:
            partes = linha.split()
            if not partes:
                continue

            # --- A. CONSUMO MEDIDO (Histórico no topo) ---
            if "Consumo Ponta" in linha and "Fora" not in linha:
                em_consumo_ponta = True
                em_consumo_fora_ponta = False
                continue

            if "Consumo Fora de Ponta" in linha:
                em_consumo_ponta = False
                em_consumo_fora_ponta = True
                continue

            if (em_consumo_ponta or em_consumo_fora_ponta) and len(partes) >= 3:
                if partes[0].isdigit() and len(partes[0]) == 4:
                    consumo = limpar_numero(partes[2])
                    injetado = limpar_numero(partes[4])

                    if em_consumo_ponta and dict_medicoes["Consumo Pt (kWh)"] == 0:
                        dict_medicoes["Consumo Pt (kWh)"] = consumo
                        dict_medicoes["Injetado Pt (kWh)"] = injetado
                        em_consumo_ponta = False

                    elif em_consumo_fora_ponta and dict_medicoes["Consumo FPt (kWh)"] == 0:
                        dict_medicoes["Consumo FPt (kWh)"] = consumo
                        dict_medicoes["Injetado FPt (kWh)"] = injetado
                        em_consumo_fora_ponta = False

        # --- C. TOTALIZADORES ---
        dict_medicoes["Consumo Total (kWh)"] = dict_medicoes["Consumo Pt (kWh)"] + dict_medicoes["Consumo FPt (kWh)"]
        dict_medicoes["Injetado Total (kWh)"] = dict_medicoes["Injetado Pt (kWh)"] + dict_medicoes["Injetado FPt (kWh)"]
        #dict_medicoes["Compensado Total (kWh)"] = dict_medicoes["Compensado Pt (kWh)"] + dict_medicoes["Compensado FPt (kWh)"]

    return dict_medicoes

def extrair_info_chave_pdf(caminho_pdf: str) -> tuple[datetime, str]:

    with pdfplumber.open(caminho_pdf) as pdf:
        pagina_um = pdf.pages[0]
        texto_pagina_um = pagina_um.extract_text()

        match_data = re.search(r"Referência:\s*(\d{2}/\d{4})", texto_pagina_um)
        if not match_data:
            raise ValueError("Não foi possível encontrar a 'Data de Referência' no PDF.")

        data_referencia = datetime.strptime(match_data.group(1), "%m/%Y")

        match_ug = re.search(r"(\d{1,3}(?:\.\d{3})+-\d{2})", texto_pagina_um)
        if not match_ug:
            raise ValueError("Não foi possível encontrar o número da 'UG' no PDF.")

        ug_valor = match_ug.group(1)

        return data_referencia, ug_valor

def dataframe_distribuidora(caminho_pdf: str) -> pd.DataFrame:

    dados_finais = []
    captura_finalizada = False

    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            if captura_finalizada:
                break

            palavras = pagina.extract_words(x_tolerance=3, y_tolerance=3)

            linhas_agrupadas = {}
            for palavra in palavras:
                pos = round(palavra['top'])
                linhas_agrupadas.setdefault(pos, []).append(palavra)

            for pos in sorted(linhas_agrupadas.keys()):
                linha_ordenada = sorted(linhas_agrupadas[pos], key=lambda p: p['x0'])
                textos_linha = [p['text'] for p in linha_ordenada]

                if any("TOTAL" in s for s in textos_linha):
                    captura_finalizada = True
                    break

                is_valida = False
                if len(textos_linha) >= 14:
                    try:
                        float(textos_linha[0])
                        int(textos_linha[1])
                        if '/' in textos_linha[-1]:
                            is_valida = True
                    except (ValueError, IndexError):
                        pass

                if is_valida:
                    dados_finais.append(textos_linha)

    if not dados_finais:
        return pd.DataFrame()

    colunas = [
        "%", "UC Comp.", "Conv GDI", "Conv GDII", "Conv GD III", "Ponta GDI",
        "Ponta GD II", "Ponta GD III", "Fora Ponta GD I", "Fora Ponta GD II",
        "Fora Ponta GD III", "Interm. GDI", "Interm. GD II", "Interm. GD III", "Ciclo"
    ]

    df = pd.DataFrame(dados_finais)

    if len(df.columns) == len(colunas):
        df.columns = colunas
    
    colunas_numericas = ["Ponta GDI", "Ponta GD II", "Ponta GD III",
                         "Fora Ponta GD I", "Fora Ponta GD II", "Fora Ponta GD III"]
    
    df[colunas_numericas] = df[colunas_numericas].apply(pd.to_numeric, errors='coerce')

    df['Recebimento'] = df["Ponta GDI"] + df["Ponta GD II"] + df["Fora Ponta GD I"] + df["Fora Ponta GD II"] + df["Ponta GD III"] + df["Fora Ponta GD III"]

    df.drop(columns=["Conv GDI", "Conv GDII", "Conv GD III",
                     "Ponta GDI", "Ponta GD II", "Ponta GD III",
                     "Fora Ponta GD I", "Fora Ponta GD II", "Fora Ponta GD III",
                     "Interm. GDI", "Interm. GD II", "Interm. GD III"], inplace=True)

    df = df[["Ciclo", "UC Comp.", "%", "Recebimento"]]

    df.rename(columns={'Ciclo':'Competencia',
                        'UC Comp.':'Instalacao',
                        '%':'Cota',
                        'Recebimento':'Creditos',},
                            inplace=True)

    df['Cota'] = df['Cota']
    df['Instalacao'] = df['Instalacao'].astype(str)
    df['Creditos'] = df['Creditos'] * -1

    return df

# ---------------------------------------------------------------------
# 2. PREPARAÇÃO DO DATAFRAME DOS RECEBIMENTOS DA AXS
# ---------------------------------------------------------------------

def copia_planilha_analise_de_faturas(diretorio_base: str) -> str:

    nome_arquivo = "Análise de Faturas.xlsm"
    caminho_original = os.path.join(diretorio_base, nome_arquivo)

    if not os.path.exists(caminho_original):
        raise FileNotFoundError(f"Arquivo 'Análise de Faturas' não encontrado: {caminho_original}")

    dir_temp = tempfile.mkdtemp()
    caminho_temporario = os.path.join(dir_temp, nome_arquivo)
    shutil.copy(caminho_original, caminho_temporario)

    return caminho_temporario

def copia_planilha_lista_vigente(diretorio_base: str) -> str:

    nome_arquivo = "Base previsão faturamento - Lista Vigente.xlsm"
    caminho_original = os.path.join(diretorio_base, nome_arquivo)

    if not os.path.exists(caminho_original):
        raise FileNotFoundError(f"Arquivo 'Lista Vigente' não encontrado: {caminho_original}")

    dir_temp = tempfile.mkdtemp()
    caminho_temporario = os.path.join(dir_temp, nome_arquivo)
    shutil.copy(caminho_original, caminho_temporario)

    return caminho_temporario

def copia_planilha_de_para(diretorio_base: str) -> str:

    nome_arquivo = "DEPARA_ENERGISA.xlsx"
    caminho_original = os.path.join(diretorio_base, nome_arquivo)

    if not os.path.exists(caminho_original):
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {caminho_original}")

    dir_temp = tempfile.mkdtemp()
    caminho_temporario = os.path.join(dir_temp, nome_arquivo)
    shutil.copy(caminho_original, caminho_temporario)

    return caminho_temporario

def dataframe_analise_de_faturas(caminho: str, data: datetime):

    global dataframe_analise

    dataframe = pd.read_excel(caminho, sheet_name="Geral", skiprows=2)

    data_ajustada = f'{str(data)[0:4]}-{int(str(data)[5:7])-1}-01'

    dataframe = dataframe[(dataframe["Mês"] == data_ajustada) & (dataframe["Distribuidora"] == "Energisa")]

    dataframe = dataframe[["Mês", "UG", "Cons P Reg", "Cons FP Reg", "Inj P Reg", "Inj FP Reg", "Saldo Anterior P kWh", "Saldo Anterior FP kWh",]]

    dataframe_analise = dataframe

    return dataframe_analise

def excedente_geracao(caminho: str, data: datetime, ug: int):

    global dataframe_tarifas
    global medicoes_axs
    global energia_recusada

    dataframe = dataframe_analise[dataframe_analise["UG"]==ug]

    dataframe_tarifas = pd.read_excel(caminho, sheet_name="Tarifas")

    data_inicial = f'{str(data)[0:4]}-{int(str(data)[5:7])-1}-01'
    
    data_final = f'{str(data)[0:4]}-{int(str(data)[5:7])-1}-{calendar.monthrange(int(str(data)[0:4]), int(str(data)[5:7])-1)[1]}'
    
    te_ponta = dataframe_tarifas[(dataframe_tarifas["Início Vigência"] <= data_final) & (dataframe_tarifas["Fim Vigência"] >= data_inicial) & (dataframe_tarifas["Sigla"] == "Energisa MT") 
                                & (dataframe_tarifas["Subgrupo"] == "A3a") & (dataframe_tarifas["Modalidade"] == "Verde") & (dataframe_tarifas["Detalhe"] == "Não se aplica") 
                                & (dataframe_tarifas["Posto"] != "Não se aplica") & (dataframe_tarifas["Unidade"] == "R$/MWh") & (dataframe_tarifas["Posto"] == "Ponta")]["TE"].values[0]

    te_fora_ponta = dataframe_tarifas[(dataframe_tarifas["Início Vigência"] <= data_final) & (dataframe_tarifas["Fim Vigência"] >= data_inicial) & (dataframe_tarifas["Sigla"] == "Energisa MT") 
                                & (dataframe_tarifas["Subgrupo"] == "A3a") & (dataframe_tarifas["Modalidade"] == "Verde") & (dataframe_tarifas["Detalhe"] == "Não se aplica") 
                                & (dataframe_tarifas["Posto"] != "Não se aplica") & (dataframe_tarifas["Unidade"] == "R$/MWh") & (dataframe_tarifas["Posto"] == "Fora ponta")]["TE"].values[0]

    coeficiente_tarifario = te_ponta/te_fora_ponta

    consumo_ponta = dataframe["Cons P Reg"].values[0]
    consumo_fora_ponta = dataframe["Cons FP Reg"].values[0]

    injetado_ponta = dataframe["Inj P Reg"].values[0]
    injetado_fora_ponta = dataframe["Inj FP Reg"].values[0]

    banco_ponta = dataframe["Saldo Anterior P kWh"].values[0]
    banco_fora_ponta = dataframe["Saldo Anterior FP kWh"].values[0]

    medicoes_axs["consumo_axs"] = round(consumo_ponta + consumo_fora_ponta)
    medicoes_axs["geracao_axs"] = round(injetado_ponta + injetado_fora_ponta)

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

    medicoes_axs["compensado_axs"] = round(credito_utilizado_ponta + credito_utilizado_fora_ponta)

    excedente = round((injetado_ponta + injetado_fora_ponta) - (credito_utilizado_ponta + credito_utilizado_fora_ponta))

    medicoes_axs["transferido_axs"] = round(excedente)

    return excedente

def dataframe_AXS(caminho: str, ug: int, data: datetime, exc) -> pd.DataFrame:

    global dataframe_analise

    data_inicial = f'{str(data)[0:4]}-{int(str(data)[5:7])-1}-01'

    dataframe_axs = pd.read_excel(caminho, sheet_name="LISTA VIGÊNTE")

    dataframe_axs = dataframe_axs[(dataframe_axs["UG"] == int(ug)) & (dataframe_axs["COMPETÊNCIA"] == data_inicial)]
    
    dataframe_axs["Qtd_recebimento"] = dataframe_axs["Quota"] * exc

    dataframe_axs["Quota"] = dataframe_axs["Quota"] * 100

    dataframe_axs["Quota"] = dataframe_axs["Quota"]

    dataframe_axs["UC"] = dataframe_axs["UC"].astype(str).str.strip()
    dataframe_axs["UC"] = dataframe_axs["UC"].str.replace(r'\.0$', '', regex=True)
    dataframe_axs["Qtd_recebimento"] = dataframe_axs["Qtd_recebimento"].astype(int)
    dataframe_axs["Qtd_recebimento"] = pd.to_numeric(dataframe_axs["Qtd_recebimento"], errors='coerce')

    return dataframe_axs

def atualizar_dataframe_AXS(dataframe: pd.DataFrame, caminho_excel: str) -> pd.DataFrame:
    
    df_axs = dataframe
    df_depara = pd.read_excel(caminho_excel, sheet_name="ENERGISA")
    
    df_depara = df_depara[["Código do cliente", "Número da UC"]]
    df_depara["Código do cliente"] = df_depara["Código do cliente"].astype(str).str.strip()
    df_depara["Número da UC"] = df_depara["Número da UC"].astype(int)
    df_depara["Número da UC"] = df_depara["Número da UC"].astype(str).str.strip()

    de_para_dict = dict(zip(df_depara["Código do cliente"], df_depara["Número da UC"]))

    df_axs["UC_Nova"] = (df_axs["UC"].map(de_para_dict).fillna(df_axs["UC"]))

    df_axs.rename(columns={'UC':'Instalacao',
                            'UC_Nova':'Instalacao_Nova',
                            'Qtd_recebimento':'Creditos',},
                                        inplace=True)
    df_axs["Instalacao_Nova"] = df_axs["Instalacao_Nova"].astype(str)

    return df_axs

# ---------------------------------------------------------------------
# 3. CONFERÊNCIAS: 1ª) MEDIÇÃO / 2ª) CRÉDITOS DE ENERGIA
# ---------------------------------------------------------------------

def validacao_medicoes() -> bool:

    if medicoes_axs["consumo_axs"] - 1 or medicoes_axs["consumo_axs"] == medicoes_distribuidora["Consumo Total (kWh)"] or medicoes_axs["consumo_axs"] + 1 == medicoes_distribuidora["Consumo Total (kWh)"] and medicoes_axs["geracao_axs"] - 1 == medicoes_distribuidora["Injetado Total (kWh)"] or medicoes_axs["geracao_axs"] == medicoes_distribuidora["Injetado Total (kWh)"] or medicoes_axs["geracao_axs"] + 1 == medicoes_distribuidora["Injetado Total (kWh)"]:

        print('\n')
        print('                  Medições Digitadas Corretas!!! \n')
        print('Consumo Digitado: {}'.format(medicoes_axs["consumo_axs"]))
        print('Consumo Real: {} \n'.format(medicoes_distribuidora["Consumo Total (kWh)"]))
        
        print('Geração Digitado: {}'.format(medicoes_axs["geracao_axs"]))
        print('Geração Real: {} \n'.format(medicoes_distribuidora["Injetado Total (kWh)"]))

        return True
    
    else:
        print('                  Medições Digitadas Incorretas!!! \n')
        print('Consumo Digitado: {}'.format(medicoes_axs["consumo_axs"]))
        print('Consumo Real: {} \n'.format(medicoes_distribuidora["Consumo Total (kWh)"]))
                
        print('Geração Digitado: {}'.format(medicoes_axs["geracao_axs"]))
        print('Geração Real: {} \n'.format(medicoes_distribuidora["Injetado Total (kWh)"]))
        
        return False

def validacao_recebimentos(dataframe_axs: pd.DataFrame, dataframe_dist: pd.DataFrame, date: datetime) -> bool:

    global dataframe_aprovados
    global dataframe_reprovados
    global dataframe_divergencias
    global dataframe_nao_esperados
    global clientes_esperados
    global clientes_utilizados
    global contador_aprovados
    global contador_reprovados

    dataframe_aprovados = pd.DataFrame(columns=['Data', 'Código', 'Cota Axs', 'Cota Distribuidora', 'Recebimento AXS', 'Recebimento Distribuidora', 'Diferenca', 'variacao'])

    dataframe_reprovados = pd.DataFrame(columns=['Data', 'Código', 'Cota Axs', 'Recebimento AXS'])

    dataframe_nao_esperados = pd.DataFrame(columns=['Data', 'Código', 'Cota Distribuidora', 'Recebimento Distribuidora'])

    for index, row in dataframe_axs.iterrows():
        data = date
        instalacao = row['Instalacao_Nova']
        instalacao_ = row['Instalacao']
        cota_axs = row["Quota"]
        recebimento_axs = row['Creditos']
        
        if instalacao in dataframe_dist['Instalacao'].values:

            cota_distribuidora = dataframe_dist[dataframe_dist['Instalacao'] == instalacao]['Cota'].values[0]
            recebimento_distribuidora = dataframe_dist[dataframe_dist['Instalacao'] == instalacao]['Creditos'].values[0]
            diferenca = recebimento_axs - recebimento_distribuidora

            if diferenca == 0:
                dataframe_aprovados.loc[len(dataframe_aprovados)] = [data, instalacao_, cota_axs, cota_distribuidora, recebimento_axs, recebimento_distribuidora, diferenca, f'0%']
            else:
                variacao = (diferenca/recebimento_axs) * 100
                dataframe_aprovados.loc[len(dataframe_aprovados)] = [data, instalacao_,  cota_axs, cota_distribuidora, recebimento_axs, recebimento_distribuidora, diferenca, f'{variacao}%']
            contador_aprovados += 1
        else:
            dataframe_reprovados.loc[len(dataframe_reprovados)] = [data, instalacao_, cota_axs, recebimento_axs]
            contador_reprovados += 1

    dataframe_divergencias = dataframe_aprovados[(dataframe_aprovados['Diferenca'] > 3) | (dataframe_aprovados['Diferenca'] < -3)]
    divergencias = dataframe_divergencias.shape[0]

    clientes_esperados = len(dataframe_axs['Instalacao_Nova'])
    clientes_utilizados = len(dataframe_dist['Instalacao'])
    
    if divergencias == 0:

        instalacoes_esperadas = set(dataframe_axs['Instalacao_Nova'])
        instalacoes_utilizadas = set(dataframe_dist['Instalacao'])

        instalacoes_nao_esperadas = instalacoes_utilizadas.difference(instalacoes_esperadas)

        for index in instalacoes_nao_esperadas:
        
            linha_instalacao = dataframe_dist[dataframe_dist['Instalacao'] == index]
            data = linha_instalacao['Competencia'].values[0]
            recebimento_distribuidora = linha_instalacao['Creditos'].values[0]
            dataframe_nao_esperados.loc[len(dataframe_nao_esperados)] = [data, index, cota_distribuidora,recebimento_distribuidora]

        return True
    else:
        return False

def ajustes_de_credito():

    global dataframe_axs

    for index, row in dataframe_divergencias.iterrows():

        codigo = row["Código"]
        credito = row["Recebimento Distribuidora"]

        dataframe_axs.loc[dataframe_axs["Instalacao_Nova"] == codigo, "Creditos"] = credito

def revalidacao_match_de_credito():

    global dataframe_aprovados
    global dataframe_reprovados

    dataframe = dataframe_nao_esperados

    lista = set(dataframe["Recebimento Distribuidora"])

    for credito in lista:
        if len(dataframe_nao_esperados["Recebimento Distribuidora"] == credito) == len(dataframe_reprovados["Recebimento AXS"] == credito):

            pergunta = input("Deseja vincular parte ou todos os clientes não encontrados aos clientes recusados ?: ")

            if pergunta == "s":
                for index, row in dataframe_reprovados[dataframe_reprovados["Recebimento AXS"] == credito].iterrows():
                    data = row["Data"]
                    codigo = row["Código"]
                    cota_axs = row["Cota Axs"]
                    cota_distribuidora = row["Cota Axs"]
                    recebimento_axs = row["Recebimento AXS"]
                    recebimento_distribuidora = row["Recebimento AXS"]
                    diferenca = 0
                    dataframe_aprovados.loc[len(dataframe_aprovados)] = [data, codigo, cota_axs, cota_distribuidora, recebimento_axs, recebimento_distribuidora, diferenca, f'0%']

                    dataframe_reprovados = dataframe_reprovados.drop(index)

                    print("Houve reinclusão de clientes recusados !!!")
            else:
                print("Não houve reinclusão de clientes recusados. A análise permanene a mesma !!!")

# ---------------------------------------------------------------------
# 4. Salvar : 1ª) DATAFRAME / 2ª) SALVAR ARQUIVO
# ---------------------------------------------------------------------

def salvar(ug: str, data: datetime, dataframe_axs: pd.DataFrame, dataframe_dist: pd.DataFrame, dataframe_ne: pd.DataFrame):

    nome_arquivo = '{}.xlsx'.format(ug)
    DIRETORIO_ARQUIVO = os.path.join(
                                    AUDITORIA_DIR,
                                    '0{}-{}'.format(data.month, data.year)
                                )

    if not os.path.exists(DIRETORIO_ARQUIVO):
        os.makedirs(DIRETORIO_ARQUIVO)

    caminho_arquivo = os.path.join(DIRETORIO_ARQUIVO, nome_arquivo)

    with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:

        dataframe_axs.to_excel(writer, sheet_name='Aprovados', index=False)
        dataframe_dist.to_excel(writer, sheet_name='Reprovados', index=False)
        dataframe_ne.to_excel(writer, sheet_name='Nao Esperados', index=False)

    print(f"Arquivo {os.path.basename(caminho_arquivo)} criado com sucesso!")

    origem = PDF_CAMINHO
    destino = os.path.join(DIRETORIO_ARQUIVO, '{}.pdf'.format(ug))

    shutil.move(origem, destino)
    
    print("Arquivo {} movido com sucesso!".format(os.path.basename(destino)))

def csv(data: datetime, ug: str ):

    dataframe = pd.DataFrame(columns=['Periodo', 'Modalidade', 'Instalacao', 'Quota', 'Posto_horario', 'Qtd_consumo', 'Qtd_geracao', 'Qtd_compensacao', 'Qtd_saldo_ant', 'Qtd_transferencia', 'Qtd_recebimento', 'Qtd_saldo_atual', 'Qtd_saldo_exp', 'Qtd_prox_saldo_exp', 'Per_prox_saldo_exp'])

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
        Instalacao = row["Código"]
        Quota = row["Cota Axs"]
        Posto_horario = "FORAP"
        Qtd_consumo = ""
        Qtd_geracao = ""
        Qtd_compensacao = ""
        Qtd_saldo_ant = ""
        Qtd_transferencia = ""
        Qtd_recebimento = row["Recebimento Distribuidora"]
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
    Qtd_consumo = medicoes_axs["consumo_axs"]
    Qtd_geracao = medicoes_axs["geracao_axs"]
    Qtd_compensacao = medicoes_axs["compensado_axs"]
    Qtd_saldo_ant = ""
    Qtd_transferencia = medicoes_axs["transferido_axs"]
    Qtd_recebimento = ""
    Qtd_saldo_atual = ""
    Qtd_saldo_exp = ""
    Qtd_prox_saldo_exp = ""
    Per_prox_saldo_exp = ""

    dataframe.loc[len(dataframe)] = [Periodo, Modalidade, Instalacao, Quota, Posto_horario, Qtd_consumo, Qtd_geracao, Qtd_compensacao, Qtd_saldo_ant, Qtd_transferencia, Qtd_recebimento, Qtd_saldo_atual, Qtd_saldo_exp, Qtd_prox_saldo_exp, Per_prox_saldo_exp]

    
    dataframe["Quota"] = dataframe["Quota"].astype(float) / 100

    nome_arquivo = '{} {}.csv'.format(str(data)[0:10], ug)

    caminho_arquivo = os.path.join('C:/Users/lucas.conceicao/Downloads', nome_arquivo)

    dataframe.to_csv(caminho_arquivo, index=False, sep=";", encoding="utf-8-sig", decimal=',')

    print(f"Arquivo CSV {os.path.basename(caminho_arquivo)} criado com sucesso!")

# ---------------------------------------------------------------------
# 0. RODAR PROGRAMA
# ---------------------------------------------------------------------
  
def principal():

    global dataframe_axs
    global dataframe_energisa
    global medicoes_distribuidora

    data, ug = extrair_info_chave_pdf(PDF_CAMINHO)

    ug_antiga = MAPA_NOVA_PARA_ANTIGA.get(ug, ug)

    medicoes_distribuidora = medicao_distribuidora(PDF_CAMINHO)
    
    dataframe_energisa = dataframe_distribuidora(PDF_CAMINHO)

    caminho_temporario = copia_planilha_analise_de_faturas(EXCEL_DIR)

    dataframe_analise_de_faturas(caminho_temporario, data)

    excedente = excedente_geracao(caminho_temporario, data, ug)

    caminho_temporario = copia_planilha_lista_vigente(EXCEL_2_DIR)

    dataframe_axs = dataframe_AXS(caminho_temporario, ug_antiga, data, excedente)
    
    caminho_temporario = copia_planilha_de_para(DEPARA_DIR)

    dataframe_axs = atualizar_dataframe_AXS(dataframe_axs, caminho_temporario)

    dataframe_axs = dataframe_axs[["COMPETÊNCIA", "Instalacao", "Instalacao_Nova", "Quota", "Creditos"]]

    if validacao_medicoes():

        def secundario():

            if validacao_recebimentos(dataframe_axs, dataframe_energisa, data):

                print("                     Todos os créditos recebidos pelos beneficiários estão corretos \n ")
                print('                                    RESUMO DOS BENEFICIÁRIOS \n')
                print('     Planilha de Recebimento  /  Demonstrativo de Compensação')
                print('  TOTAL                  {}      {}'.format(clientes_esperados, clientes_utilizados))
                print('APROVADOS                {}      {}'.format(contador_aprovados, contador_aprovados))
                print('REPROVADOS               {} \n'.format(contador_reprovados))
                print('                          ************  APROVADOS  ************')
                colunas = dataframe_aprovados.columns.tolist()
                print(tabulate(dataframe_aprovados, headers=colunas, tablefmt='fancy_grid'))
                print(' ')
                print('                 ***********  REPROVADOS  ************')
                colunas = dataframe_reprovados.columns.tolist()
                print(tabulate(dataframe_reprovados, headers=colunas, tablefmt='fancy_grid'))
                print(' ')
                print('                 ***********  NÃO ESPERADOS  ************')
                colunas = dataframe_nao_esperados.columns.tolist()
                print(tabulate(dataframe_nao_esperados, headers=colunas, tablefmt='fancy_grid'))

                if len(dataframe_nao_esperados) != 0:
                    revalidacao_match_de_credito()

                pergunta = input("Deseja realizar a função de criação e salvamento do arquivo de conferência na pasta de destino: ")

                if pergunta.lower() == "s":
                    salvar(ug_antiga, data, dataframe_aprovados, dataframe_reprovados, dataframe_nao_esperados)
                    csv(data, ug_antiga)
                else:
                    print("Os arquivo não foi criado")
            else:
                print('                                          Há divergências de alguns créditos \n \n ')
                print('                                               RESUMO DOS BENEFICIÁRIOS \n \n ')
                colunas = dataframe_divergencias.columns.tolist()
                print(tabulate(dataframe_divergencias, headers=colunas, tablefmt='fancy_grid'))

                pergunta = input("Deseja realizar a função de ajuste dos créditos: ")
                        
                if pergunta.lower() == "s":
                    ajustes_de_credito()
                    secundario()
                else:
                    print("Não foi criado o arquivo de conferência")

        secundario()
        

principal()