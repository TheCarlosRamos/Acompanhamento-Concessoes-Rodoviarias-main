# %% [markdown]
# <h1><b><i>SCRIPT PRINCIPAL</b></i></h1>
# <h2><b><i>Importação de Bibliotécas</b></i></h2>

# %%
import sys
sys.path.append(r'C:\pylibs')

import pandas as pd
import numpy as np
import openpyxl

# %%
import pandas as pd
import openpyxl
import numpy as np
import os

# %%
# ==============================================================================
# 1. SETUP INICIAL E CARREGAMENTO
#    (Melhor prática: definir constantes e carregar o arquivo de forma segura)
# ==============================================================================

# Definição do caminho do arquivo em uma variável para fácil manutenção
FILE_PATH = 'Planilha de Monitoramento_PPI_2025_31_03_25_PPI_R21.xlsx'
OUTPUT_DIR = 'Dados Gerados'

# Verifica se o arquivo existe antes de tentar carregá-lo
if not os.path.exists(FILE_PATH):
    print(f"ERRO: O arquivo '{FILE_PATH}' não foi encontrado.")
    # Encerra o script ou lança uma exceção se o arquivo principal não existir
    exit()

# Cria a pasta de saída se ela não existir
os.makedirs(OUTPUT_DIR, exist_ok=True)

print(f"Arquivo '{FILE_PATH}' encontrado. Iniciando processamento...")

# %%
# Listar todas as planilhas disponíveis no arquivo
caminho_arquivo = "Planilha Monitoramento_PPI - Carregamento - BI 2025 - 7.xlsx"
xls = pd.ExcelFile(caminho_arquivo)
sheet_names = xls.sheet_names
sheet_names

# %%
file_path = 'Planilha de Monitoramento_PPI_2025_31_03_25_PPI_R21 - informacao.xlsx'

# %%
# Carregar as planilhas

df_informacoes = pd.read_excel(file_path, sheet_name='INFORMAÇÕES (1)')
df_per = pd.read_excel(file_path, sheet_name='PER (1)')
df_plan_exec_ano_ant = pd.read_excel(file_path, sheet_name='PLAN_EXEC ATÉ ANO ANTERIOR (22)')
df_meta_23 = pd.read_excel(file_path, sheet_name='META(1) 2023')
df_meta_24 = pd.read_excel(file_path, sheet_name='META(2024)')
df_meta_25 = pd.read_excel(file_path, sheet_name='META(2025)')
df_meta_26 = pd.read_excel(file_path, sheet_name='META(2026)')
df_meta_27 = pd.read_excel(file_path, sheet_name='META(2027)')
df_inexec = pd.read_excel(file_path, sheet_name='iNEXECUTADO (pendências até 24)')
df_a_executar = pd.read_excel(file_path, sheet_name='À EXECUTAR (2026 DIANTE)')

# %% [markdown]
# <h1><b><i>INFORMAÇÕES (1)</b></i></h1>

# %%
#CHECKPOINT
df_informacoes_copy = df_informacoes.copy()

# %%
df_informacoes_copy = pd.read_excel(file_path, sheet_name='INFORMAÇÕES (1)', header=2)
df_informacoes_copy.columns

# %%
df_informacoes_copy

# %%
df_informacoes_copy = df_informacoes_copy.drop([0, 1, 2], axis = 0)

# %%
df_informacoes_copy = df_informacoes_copy.dropna(subset=['BR'])

# %%
df_informacoes_copy['Etapa'] = None
df_informacoes_copy['Descrição do Trecho'] = None
etapa_atual = None

for index, row in df_informacoes_copy.iterrows():
    if 'ETAPA' in str(row['TRECHO']):
        etapa_atual = row['TRECHO']
    df_informacoes_copy.at[index, 'Etapa'] = etapa_atual
    if pd.notna(row['TRECHO']) and 'ETAPA' not in str(row['TRECHO']):
        df_informacoes_copy.at[index, 'Descrição do Trecho'] = row['TRECHO']

# %%
df_informacoes_copy.drop('TRECHO', axis=1, inplace=True)

# %%
# Updating the 'Região' column based on the 'UF' mapping
region_mapping = {
    'AC': 'Norte', 'AP': 'Norte', 'AM': 'Norte', 'PA': 'Norte', 'RO': 'Norte', 'RR': 'Norte', 'TO': 'Norte',
    'AL': 'Nordeste', 'BA': 'Nordeste', 'CE': 'Nordeste', 'MA': 'Nordeste', 'PB': 'Nordeste', 'PE': 'Nordeste',
    'PI': 'Nordeste', 'RN': 'Nordeste', 'SE': 'Nordeste',
    'DF': 'Centro-Oeste', 'GO': 'Centro-Oeste', 'MS': 'Centro-Oeste', 'MT': 'Centro-Oeste',
    'ES': 'Sudeste', 'MG': 'Sudeste', 'RJ': 'Sudeste', 'SP': 'Sudeste',
    'PR': 'Sul', 'SC': 'Sul', 'RS': 'Sul'
}
df_informacoes_copy['Região'] = df_informacoes_copy['UF'].map(region_mapping)

# %%
columns_to_fill = ['ESTADO/LOTE', 'EMPREENDIMENTO', 'PROPONENTE', 'EXECUTOR             (Grupo Controlador)',
       'ESTRUTURADOR DO PROJETO', 'ANO LEILÃO', 'DATA DE INÍCIO']
df_informacoes_copy[columns_to_fill] = df_informacoes_copy.groupby((df_informacoes_copy['BR'] == 1).cumsum())[columns_to_fill].transform(lambda x: x.ffill())

# %%
df_informacoes_copy['CAPEX (BI)'].unique()

# %%
df_informacoes_copy

# %%
import pandas as pd
import numpy as np

# 1) Identificar as linhas onde 'KM Inicial' tem o texto de situação
situacoes_validas = ['ATIVO', 'CADUCIDADE', 'JUDICIAL', 'RELICITAÇÃO', 'RENEGOCIAÇÃO']

# Cria 'Situação' com os rótulos onde aparecerem e NaN no resto
df_informacoes_copy['Situação'] = df_informacoes_copy['KM Inicial'].apply(
    lambda x: str(x).strip().upper() if str(x).strip().upper() in situacoes_validas else pd.NA
)

# 2) Preencher para baixo (forward fill) para propagar a situação
df_informacoes_copy['Situação'] = df_informacoes_copy['Situação'].ffill()

# 3) Extrair valor numérico de 'KM Inicial' para a coluna 'km (i)'
#    - Converte tudo para string
#    - Troca vírgula por ponto (estilo BR -> US)
#    - Remove espaços
#    - Converte para número, jogando não numéricos para NaN
km_str = (
    df_informacoes_copy['KM Inicial']
    .astype(str)
    .str.replace('.', '', regex=False)        # se existir separador de milhar como ponto (ex.: 1.234,56)
    .str.replace(',', '.', regex=False)       # vírgula decimal -> ponto
    .str.strip()
)

df_informacoes_copy['km (i)'] = pd.to_numeric(km_str, errors='coerce')

# (Opcional) Se quiser garantir que quando a célula é uma Situação, o km (i) vire NaN:
mask_situacao = df_informacoes_copy['KM Inicial'].astype(str).str.strip().str.upper().isin(situacoes_validas)
df_informacoes_copy.loc[mask_situacao, 'km (i)'] = pd.NA

# %%
df_informacoes_copy.drop('KM Inicial', axis=1, inplace=True)

# %%
# Renomear Coluna ['EXECUTOR...']
# DROP de Colunas e Realocação.

df_informacoes_copy.rename(columns={'EXECUTOR             (Grupo Controlador)': 'EXECUTOR (Grupo Controlador)'}, inplace=True)
df_informacoes_copy.rename(columns={'EXTENSÃO\n (km)': 'EXTENSÃO (km)'}, inplace=True)
df_informacoes_copy.rename(columns={'KM Final': 'km (f)'}, inplace=True)

df_informacoes_copy = df_informacoes_copy[['Região', 'SETOR', 'UF', 'ESTADO/LOTE', 'BR', 'EMPREENDIMENTO', 'PROPONENTE',
       'EXECUTOR (Grupo Controlador)', 'ESTRUTURADOR DO PROJETO',
       'ANO LEILÃO', 'DATA DE INÍCIO', 'ANO DA CONCESSÃO',
       'PRAZO (anos)', 'CAPEX (BI)', 'OPEX (BI)', 'INVESTIMENTO TOTAL (BI)', 
       'Etapa', 'Descrição do Trecho',
       'Situação', 'km (i)', 'km (f)', 'EXTENSÃO (km)']]

# %%
df_informacoes_copy['ESTRUTURADOR DO PROJETO'].unique()

# %%
df_informacoes_copy['PROPONENTE'].unique()

# %%
df_informacoes_copy.columns

# %%
# Converter colunas para 'category'
df_informacoes_copy['SETOR'] = df_informacoes_copy['SETOR'].astype('category')
df_informacoes_copy['UF'] = df_informacoes_copy['UF'].astype('category')
df_informacoes_copy['ESTADO/LOTE'] = df_informacoes_copy['ESTADO/LOTE'].astype('category')
df_informacoes_copy['PROPONENTE'] = df_informacoes_copy['PROPONENTE'].astype('category')
df_informacoes_copy['EXECUTOR (Grupo Controlador)'] = df_informacoes_copy['EXECUTOR (Grupo Controlador)'].astype('category')
df_informacoes_copy['ESTRUTURADOR DO PROJETO'] = df_informacoes_copy['ESTRUTURADOR DO PROJETO'].astype('category')

# Converter colunas para 'int'
df_informacoes_copy['ANO DA CONCESSÃO'] = pd.to_numeric(df_informacoes_copy['ANO DA CONCESSÃO'], errors='coerce').fillna(0).astype(int)
df_informacoes_copy['PRAZO (anos)'] = pd.to_numeric(df_informacoes_copy['PRAZO (anos)'], errors='coerce').fillna(0).astype(int)

# Converter colunas para 'float'
df_informacoes_copy['CAPEX (BI)'] = pd.to_numeric(df_informacoes_copy['CAPEX (BI)'], errors='coerce')
df_informacoes_copy['OPEX (BI)'] = pd.to_numeric(df_informacoes_copy['OPEX (BI)'], errors='coerce')
df_informacoes_copy['INVESTIMENTO TOTAL (BI)'] = df_informacoes_copy['INVESTIMENTO TOTAL (BI)'].replace('[\$]', '', regex=True).replace(',', '.', regex=True)
df_informacoes_copy['INVESTIMENTO TOTAL (BI)'] = pd.to_numeric(df_informacoes_copy['INVESTIMENTO TOTAL (BI)'], errors='coerce')

df_informacoes_copy['km (i)'] = pd.to_numeric(df_informacoes_copy['km (i)'], errors='coerce')
df_informacoes_copy['km (f)'] = pd.to_numeric(df_informacoes_copy['km (f)'], errors='coerce')
df_informacoes_copy['EXTENSÃO (km)'] = pd.to_numeric(df_informacoes_copy['EXTENSÃO (km)'], errors='coerce')

# Converter colunas para 'string'
df_informacoes_copy['EMPREENDIMENTO'] = df_informacoes_copy['EMPREENDIMENTO'].astype('string') 


# %%
# Função para limpar e converter as datas para o formato datetime
def limpar_e_converter_data(data):
    # Tratar dados como string para evitar problemas com diferentes tipos
    data = str(data)
    # Substituir '//'' por '/'
    data = data.replace('//', '/')
    # Substituir '/' por '-'
    data = data.replace('/', '-')
    # Tentar converter para datetime
    return pd.to_datetime(data, errors='coerce', dayfirst=True)

# Aplicar a função no DataFrame
# Substitua 'df_informacoes_copy' pelo nome do seu DataFrame e 'DATA DE INÍCIO' pelo nome da coluna que contém as datas
df_informacoes_copy['DATA DE INÍCIO'] = df_informacoes_copy['DATA DE INÍCIO'].apply(limpar_e_converter_data)

# %%
columns_to_fill = ['Região', 'ESTADO/LOTE', 'EMPREENDIMENTO', 'EXECUTOR (Grupo Controlador)']
df_informacoes_copy[columns_to_fill] = df_informacoes_copy.groupby((df_informacoes_copy['BR'] == 1).cumsum())[columns_to_fill].transform(lambda x: x.ffill())

# %%
df_informacoes_copy.index = range(1, len(df_informacoes_copy) + 1)

# %%
df_informacoes_copy.index.name = 'ID-ÚNICO'

# %%
df_informacoes_copy

# %%
df_informacoes_copy['UF'].unique()


