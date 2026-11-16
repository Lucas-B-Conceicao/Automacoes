# Tentar extrair apenas o texto da segunda página em diante para entender a estrutura
import pdfplumber
texto_paginas = []
caminho_arquivo = "C:/Users/lucas.conceicao/Downloads/download.pdf"
with pdfplumber.open(caminho_arquivo) as pdf:
    for page in pdf.pages[1:]:  # Pega as páginas a partir da segunda
        texto_paginas.append(page.extract_text())

# Exibir uma amostra do texto extraído para análise
print(texto_paginas[:3])  # Mostra as 3 primeiras páginas analisadas a partir da segunda página

import re
import pandas as pd

# Inicializar uma lista para armazenar as linhas da tabela
linhas_tabela = []

# Percorrer o texto extraído e capturar apenas as linhas relevantes da tabela
capturando = False

for pagina in texto_paginas:
    for linha in pagina.split("\n"):
        if "Discriminação das Transferências do Mês" in linha:
            capturando = True  # Ativa a captura ao encontrar o título da tabela
            continue

        if capturando:
            if re.match(r"^\d{4,}", linha):  # Linhas que começam com um número grande (provável UC)
                linhas_tabela.append(linha)

# Separar os dados por espaços e estruturar em um DataFrame
dados_formatados = [re.split(r'\s+', linha.strip()) for linha in linhas_tabela]

# Criar DataFrame com colunas nomeadas
colunas = ["UC", "Comp. Conv", "Ponta", "Fora Ponta", "Interm.", "Ciclo"]
df_transferencias = pd.DataFrame(dados_formatados, columns=colunas)

# Exibir a tabela organizada
print(df_transferencias)
