from pathlib import Path
import pandas as pd

# 1. Definir o caminho da pasta
caminho_pasta = Path(r"C:\Users\lucas.conceicao\Relatorio de Compensacao\relatorios")

# 2. Criar uma lista para guardar os dados de cada planilha
lista_tabelas = []

print("Buscando planilhas...")

# 3. Varrer a pasta procurando por arquivos .xlsx ou .xls
# (O método 'glob' filtra apenas arquivos do Excel)
for arquivo in caminho_pasta.glob("*.xlsx"):
    print(f"Lendo: {arquivo.name}")
    
    # Lê a planilha atual
    df = pd.read_excel(arquivo, dtype={"Novo Código": str})
    
    # Opcional: Adiciona uma coluna para você saber de qual arquivo veio cada linha
    df["Origem_Arquivo"] = arquivo.name
    
    # Guarda os dados na nossa lista
    lista_tabelas.append(df)

# 4. Juntar tudo se houver arquivos encontrados
if lista_tabelas:
    print("\nMesclando todas as planilhas...")
    # 'concat' junta todos os DataFrames da lista verticalmente (um embaixo do outro)
    consolidado = pd.concat(lista_tabelas, ignore_index=True)
    
    # Definir o caminho de saída (salva na mesma pasta com o nome 'consolidado.xlsx')
    caminho_saida = caminho_pasta / "consolidado_final.xlsx"
    
    # Salvar o resultado final
    # 'index=False' evita que o pandas crie uma coluna extra de números de linha
    consolidado.to_excel(caminho_saida, index=False)
    print(f"Sucesso! Planilha única salva em: {caminho_saida}")
else:
    print("Nenhum arquivo .xlsx foi encontrado na pasta especificada.")