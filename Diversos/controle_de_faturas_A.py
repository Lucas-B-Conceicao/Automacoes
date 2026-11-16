import os
import pandas as pd
import pyodbc

class InclusaoEmMassa:

    def __init__(self):

        self.PASTA_USUARIO = os.path.expanduser('~')
        
        self.DIRETORIO_RAIZ = os.path.join(
            self.PASTA_USUARIO,
            "OneDrive 2026",
            "AXS ENERGIA S A", 
            "Gestão de Usinas - Usinas AXS"
            )
        
        self.SUB_DIRETORIO = os.path.join(
                self.PASTA_USUARIO,
                self.DIRETORIO_RAIZ,
                "03.Análises",
                "14 - Controle de Faturas e Recebimentos",
                        )
        
        self.caminho_banco = os.path.join(
            self.SUB_DIRETORIO,
        'ControleFaturasEnergia.accdb'
        )

        self.string_conexao = (
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={self.caminho_banco};"
        )

        self.conexao = None

    def __enter__(self):

        try:
            self.conexao = pyodbc.connect(self.string_conexao)
            print("Conexão estabelecida com sucesso!")

            return self.conexao.cursor()
        
        except pyodbc.Error as erro:
            nome_arquivo = os.path.basename(self.caminho_banco)
            print(f"Erro ao conetar ao banco de dados {nome_arquivo}: {erro}")
            raise

    def __exit__(self, exc_type, exc, tb):

        if self.conexao:
            if exc_type is None:
                self.conexao.commit()
            else:
                self.conexao.rollback()

            self.conexao.close

banco_de_dados = InclusaoEmMassa()

with banco_de_dados as cursor:

    caminho_UGs_Dist = os.path.join(
                    banco_de_dados.SUB_DIRETORIO,
                    "UGs x Dist.xlsm",
                )
    
    dataframe = pd.read_excel(caminho_UGs_Dist, sheet_name="Recebimento e Fatura")

    dataframe = dataframe[["Referência", "CRI", "Distribuidora", "USINA", "UG", "Lista vigente?", "Emissão dia", "Baixadas", "Lançadas", "Análise de Fatura", "Recebimento?", "NO ERP?"]]
    dataframe = dataframe[dataframe["Referência"] < "2026-11-01 "]
    print(dataframe)
    dataframe["CRI"] = dataframe["CRI"].astype(str)
    dataframe["UG"] = dataframe["UG"].astype(str)
 
    for index, row in dataframe.iterrows():

        competencia = row["Referência"]
        consorcio = row["CRI"]
        distribuidora = row["Distribuidora"]
        usina = row["USINA"]
        unidadegeradora = f'{row["UG"]}'
        listavigente = row["Lista vigente?"]
        diadeemissao = row["Emissão dia"]
        baixada = row["Baixadas"]
        enviada = row["Lançadas"]
        auditada = row["Análise de Fatura"]
        recebimento = row["Recebimento?"]
        erp = row["NO ERP?"]

        consulta = "INSERT INTO ControledeFaturasdeEnergia (Competencia, Consorcio, Distribuidora, Usina, UnidadeGeradora, ListaVigente, DiadeEmissao, Baixada, Enviada, Auditada, Recebimento, ERP) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"

        cursor.execute(consulta, (competencia, consorcio, distribuidora, usina, unidadegeradora, listavigente, diadeemissao, baixada, enviada, auditada, recebimento, erp))
