import os
import pandas as pd
import pyodbc

dicionario_usinas = {3014292343:"Paulo Valias", 3014292356:"Paulo Valias", 3014292382:"Paulo Valias",
              3014109500:"Boa Vista I", 3014109501:"Boa Vista I", 3014109502:"Boa Vista I", 
              3014233705:"Harmonia I", 3014926608:"Harmonia I", 3014926612:"Harmonia I",
              3014298175:"Harmonia II", 3014298176:"Harmonia II",
              3014233708:"Boa Vista II", 3015083428:"Boa Vista II", 3015083492:"Boa Vista II",
              3014427046:"Itatiaiucu", 3015120874:"Itatiaiucu",
              3014583209:"Prata I",3014583260:"Prata I",3014583262:"Prata I",
              3014837599:"Frei Inocencio I", 3014837605:"Frei Inocencio I", 3014837606:"Frei Inocencio I",
              3014306059:"Claudio", 3015948206:"Claudio",
              3014902567:"Guaxupe", 3014902579:"Guaxupe", 3014902584:"Guaxupe",
              3014868833:"Tres Coracoes II", 3014868860:"Tres Coracoes II", 3014868836:"Tres Coracoes II",
              3014900403:"Tres Pontas I", 3014900405:"Tres Pontas I", 3014900389:"Tres Pontas I",
              3014629614:"Campina Verde I", 3014629582:"Campina Verde I", 3014629309:"Campina Verde I", 3014629370:"Campina Verde I",
              3014784085:"Prata II", 3014784110:"Prata II", 3014784111:"Prata II",
              3014965842:"Padre Paraiso I", 3014965846:"Padre Paraiso I", 3014965872:"Padre Paraiso I",
              3014868655:"Santa Rita de Caldas I",3014868693:"Santa Rita de Caldas I",
              3015056747:"Ouro Fino I",
              "1.962.316.017-60":"Campo Verde I","871.084.017-42":"Campo Verde I","1.962.349.017-81":"Campo Verde I",
              "1.941.939.017-62":"Tangara da Serra I","1.941.950.017-07":"Tangara da Serra I","1.941.958.017-30":"Tangara da Serra I",
              "1.941.938.017-77":"Tangara da Serra II","1.941.940.017-25":"Tangara da Serra II","1.941.968.017-12":"Tangara da Serra II",
              "1.102.319.017-42":"Juruena II","1.156.316.017-01":"Juruena II","979.296.017-05":"Juruena II",
              "362.232.017-86":"Claudia I","2.001.766.017-02":"Claudia I","2.001.770.017-39":"Claudia I",
              "955.455.017-79":"Brasnorte I","275.546.017-41":"Brasnorte I","1.983.148.017-16":"Brasnorte I","1.455.682.017-04":"Nova Ubirata I","2.000.159.017-90":"Nova Ubirata I","1.237.069.017-60":"Nova Ubirata I","1.982.149.017-47":"Araputanga I","1.314.586.017-60":"Araputanga I","1.192.318.017-45":"Araputanga I","2.070.805.017-03":"Feliz Natal I","2.054.945.017-55":"Feliz Natal I","1.416.330.017-38":"Feliz Natal I",
              "667.973.017-60":"Paranaita I","882.913.017-53":"Paranaita I","1.404.417.017-03":"Paranaita I","2.090.419.017-76":"Paranaita II","2.062.197.017-91":"Paranaita II","2.062.397.017-38":"Paranaita II","815.409.017-60":"Paranaita III","382.776.017-48":"Paranaita III","382.777.017-33":"Paranaita III","18.890.017-09":"Nova Ubirata II","1.224.855.017-57":"Nova Ubirata II","2.016.632.017-98":"Nova Ubirata II"}


dicionario_consorcios = {"Paulo Valias":"Consorcio 01","Harmonia I":"Consorcio 01","Harmonia II":"Consorcio 01","Boa Vista I":"Consorcio 01","Boa Vista II":"Consorcio 01","Itatiaiucu":"Consorcio 01","Prata I":"Consorcio 01",
                         "Frei Inocencio I":"Consorcio 02","Claudio":"Consorcio 02",
                         "Guaxupe":"Consorcio 03","Tres Coracoes II":"Consorcio 03","Tres Pontas I":"Consorcio 03","Campina Verde I":"Consorcio 03","Prata II":"Consorcio 03","Campo Verde I":"Consorcio 03",
                         "Padre Paraiso I":"Consorcio 04","Tangara da Serra II":"Consorcio 04","Tangara da Serra I":"Consorcio 04",
                         "Juruena II":"Consorcio 05",
                         "Santa Rita de Caldas I":"Consorcio 06","Claudia I":"Consorcio 06",
                         "Brasnorte I":"Consorcio 07","Nova Ubirata I":"Consorcio 07","Araputanga I":"Consorcio 07","Feliz Natal I":"Consorcio 07",
                         "Nova Ubirata II":"Consorcio 08","Paranaita III":"Consorcio 08","Paranaita II":"Consorcio 08","Paranaita I":"Consorcio 08",
                         "Ouro Fino I":"Consorcio 09"
}

class ProjetoElefante:

    def __init__(self):

        self.PASTA_USUARIO = os.path.expanduser('~')
        
        self.DIRETORIO_RAIZ = os.path.join(
            self.PASTA_USUARIO,
            "OneDrive 2026",
            "AXS ENERGIA S A", 
            "Gestão de Usinas - Usinas AXS"
            )
        
        self.DIRETORIO_RELATORIOS = os.path.join(
                self.PASTA_USUARIO,
                self.DIRETORIO_RAIZ,
                "03.Análises",
                "14 - Controle de Faturas e Recebimentos",
                "Recebimento",
                "Relatórios de Recebimentos",
                        )
            
        self.DIRETORIO_BANCO = os.path.join(
                self.DIRETORIO_RAIZ,
                self.PASTA_USUARIO,
                self.DIRETORIO_RAIZ,
                "03.Análises",
                "14 - Controle de Faturas e Recebimentos",
                "Recebimento",
                "Relatórios de Recebimentos",
            )
        
        self.caminho_banco = os.path.join(
            self.DIRETORIO_BANCO,
        'ProjetoElefante.accdb'
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

banco_de_dados = ProjetoElefante()

with banco_de_dados as cursor:

    distribuidora = input("Digite a distribuidora: ")
    competencia = input("Digite a competencia (01-2022): ")

    if distribuidora in ["Cemig","ENERGISA"]:

        caminho_pasta = os.path.join(
                        banco_de_dados.DIRETORIO_RELATORIOS,
                        distribuidora,
                        competencia
                    )
        
        lista_pastas_uc = os.listdir(caminho_pasta)
        
        competencia = f'01/{competencia[0:2]}/{competencia[3:7]}'
        
        if distribuidora == "Cemig":
    
            for pasta_uc in lista_pastas_uc:
    
                usina = dicionario_usinas[int(pasta_uc)]
                consorcio = dicionario_consorcios[usina]
                unidadegeradora = str(pasta_uc)
    
                print(usina)
                print(consorcio)
                print(unidadegeradora)
                
                dataframe = pd.read_excel(os.path.join(caminho_pasta, pasta_uc, "{}c.xlsx".format(pasta_uc)), sheet_name="Aprovados")
    
                if 901758101883 in dataframe["Código"].values:
    
                    credito = dataframe[dataframe["Código"] == 901758101883]["Recebimento Distribuidora"].values[0]
    
                else:
    
                    credito = 0
    
                consulta = "INSERT INTO UC_Ancora_cemig (Competencia,Usina,UnidadeGeradora,CreditosdeEnergia,Consorcio) VALUES (?,?,?,?,?)"
    
                cursor.execute(consulta, (competencia,usina,unidadegeradora,float(credito),consorcio))
        
        if distribuidora == "ENERGISA":
    
            for pasta_uc in lista_pastas_uc:
    
                usina = dicionario_usinas[pasta_uc[:-5]]
                consorcio = dicionario_consorcios[usina]
                unidadegeradora = str(pasta_uc[:-5])
    
                print(usina)
                print(consorcio)
                print(unidadegeradora)
                
                dataframe = pd.read_excel(os.path.join(caminho_pasta, pasta_uc), sheet_name="Aprovados")
    
                if 54483901765 in dataframe["Código"].values:
    
                    credito = dataframe[dataframe["Código"] == 54483901765]["Recebimento Distribuidora"].values[0]
    
                else:
    
                    credito = 0
    
                consulta = "INSERT INTO UC_Ancora_Energisa (Competencia,Usina,UnidadeGeradora,CreditosdeEnergia,Consorcio) VALUES (?,?,?,?,?)"
                
                cursor.execute(consulta, (competencia,usina,unidadegeradora,float(credito),consorcio))

    else:
        print("Distribuidora não existe / Erro de Digitação")

    

    