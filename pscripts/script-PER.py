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
file_path = 'Planilha de Monitoramento_PPI_2025_31_03_25_PPI_R21.xlsx'

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

# %% [markdown]
# ---

# %%
df_informacoes_copy = df_informacoes_copy[df_informacoes_copy['BR'] == 1]

# %%
columns_to_drop = ['SETOR', 'Região', 'Descrição do Trecho', 'km (i)', 'km (f)']

# Removendo as colunas extras do DataFrame df_informacoes_copy
df_informacoes_copy = df_informacoes_copy.drop(columns=columns_to_drop)

# %%
# Dicionário de renomeação
rename_dict = {
    'UF': 'Região',
    'Etapa': 'ETAPA',
    'Situação': 'SITUAÇÃO'
}

# Renomear as colunas usando o dicionário
df_informacoes_copy = df_informacoes_copy.rename(columns=rename_dict)

# %%
df_informacoes_copy = df_informacoes_copy[['Região', 'BR', 'EMPREENDIMENTO', 'PROPONENTE',
       'EXECUTOR (Grupo Controlador)', 'ESTRUTURADOR DO PROJETO',
       'ANO LEILÃO', 'DATA DE INÍCIO', 'ANO DA CONCESSÃO',
       'PRAZO (anos)', 'CAPEX (BI)', 'OPEX (BI)', 'INVESTIMENTO TOTAL (BI)',
       'ETAPA', 'SITUAÇÃO', 'EXTENSÃO (km)', 'ESTADO/LOTE']]

# %%
# Realiza a substituição na coluna 'EMPREENDIMENTO'
df_informacoes_copy['EMPREENDIMENTO'] = df_informacoes_copy['EMPREENDIMENTO'].replace('CONCESSÃO VIA SUL', 'VIA SUL')

# %%
df_informacoes_copy.to_excel('Dados Gerados/INFORMAÇÕES (1).xlsx', index=False, sheet_name='INFORMAÇÕES (1)')

# %% [markdown]
# <h1><b><i>PER (1)</b></i></h1>

# %%
#CHECKPOINT
df_per_copy = df_per.copy()

# %%
df_per_copy = pd.read_excel(file_path, sheet_name='PER (1)', header=2)

# %%
df_per_copy

# %%
df_per_copy = df_per_copy.drop([0, 1, 2], axis = 0)

# %%
df_per_copy = df_per_copy.loc[df_per_copy['BR'] != 1]
df_per_copy = df_per_copy.dropna(subset=['BR'])

# %%
df_per_copy = df_per_copy.dropna(subset=['ESTADO/LOTE'])

# %%
import pandas as pd

# Lista de estados válidos, normalizada sem acentos
estados_validos = {
    'RIO GRANDE DO SUL', 'MINAS GERAIS', 'GOIAS', 'SANTA CATARINA', 'TOCANTINS',
    'RIO DE JANEIRO', 'MATO GROSSO', 'PARA', 'SAO PAULO', 'PARANA', 'BAHIA',
    'ESPIRITO SANTO', 'MATO GROSSO DO SUL'
}

# 1) Normaliza a coluna para comparar (tira acentos, upper, trim)
s_norm = (
    df_per_copy['ESTADO/LOTE']
    .astype(str)
    .str.strip()
    .str.upper()
    .str.normalize('NFD')               # separa acentos
    .str.encode('ascii', 'ignore')      # remove acentos
    .str.decode('utf-8')                # volta para string normal
)

# 2) Mantém apenas os valores que são estados, o resto vira NA
mask = s_norm.isin(estados_validos)
df_per_copy['ESTADO/LOTE'] = df_per_copy['ESTADO/LOTE'].where(mask, pd.NA)

# 3) Preenche para baixo com forward fill (forma compatível)
df_per_copy['ESTADO/LOTE'] = df_per_copy['ESTADO/LOTE'].ffill()

# %%
df_per_copy

# %%
df_per_copy = df_per_copy[~df_per_copy.apply(lambda row: row.astype(str).str.contains('TOTAL|SOMA').any(), axis=1)]

# %%
# Renomear Coluna ['EXECUTOR...']

df_per_copy.rename(columns={'EXECUTOR             (Grupo Controlador)': 'EXECUTOR (Grupo Controlador)'}, inplace=True)
df_per_copy.rename(columns={'PAVIMENTAÇÃO': 'PAVIMENTAÇÃO - < 2022 - Descrição'}, inplace=True)
df_per_copy.rename(columns={'DUPLICAÇÃO': 'DUPLICAÇÃO - < 2022 - Descrição'}, inplace=True)
df_per_copy.rename(columns={'OAE': 'OAE - < 2022 - Descrição'}, inplace=True)
df_per_copy.rename(columns={'CONTORNO': 'CONTORNO - < 2022 - Descrição'}, inplace=True)
df_per_copy.rename(columns={'FX ADICIONAL': 'FX ADICIONAL - < 2022 - Descrição'}, inplace=True)
df_per_copy.rename(columns={'TERCEIRA FAIXA': 'TERCEIRA FAIXA - < 2022 - Descrição'}, inplace=True)

# %%
df_per_copy

# %%
new_column_names = ['PAVIMENTAÇÃO - < 2022 - km (i)', 
                    'PAVIMENTAÇÃO - < 2022 - km (f)', 'PAVIMENTAÇÃO - < 2022 - Ext. (km)', 
                    'PAVIMENTAÇÃO - < 2022 - PERCENTUAL (%)', 'PAVIMENTAÇÃO - < 2022 - (km)%', 
                    'PAVIMENTAÇÃO - < 2022 - FINANCEIRO (R$)', 'PAVIMENTAÇÃO - 2023 - Descrição', 
                    'PAVIMENTAÇÃO - 2023 - km (i)', 'PAVIMENTAÇÃO - 2023 - km (f)', 'PAVIMENTAÇÃO - 2023 - Ext. (km)', 
                    'PAVIMENTAÇÃO - 2023 - PERCENTUAL (%)', 'PAVIMENTAÇÃO - 2023 - (km)%', 'PAVIMENTAÇÃO - 2023 - FINANCEIRO (R$)', 
                    'PAVIMENTAÇÃO - 2024 - Descrição', 'PAVIMENTAÇÃO - 2024 - km (i)', 'PAVIMENTAÇÃO - 2024 - km (f)', 
                    'PAVIMENTAÇÃO - 2024 - Ext. (km)', 'PAVIMENTAÇÃO - 2024 - PERCENTUAL (%)', 'PAVIMENTAÇÃO - 2024 - (km)%', 
                    'PAVIMENTAÇÃO - 2024 - FINANCEIRO (R$)', 'PAVIMENTAÇÃO - 2025 - Descrição', 'PAVIMENTAÇÃO - 2025 - km (i)', 
                    'PAVIMENTAÇÃO - 2025 - km (f)', 'PAVIMENTAÇÃO - 2025 - Ext. (km)', 'PAVIMENTAÇÃO - 2025 - PERCENTUAL (%)', 
                    'PAVIMENTAÇÃO - 2025 - (km)%', 'PAVIMENTAÇÃO - 2025 - FINANCEIRO (R$)', 'PAVIMENTAÇÃO - 2026 - Descrição', 
                    'PAVIMENTAÇÃO - 2026 - km (i)', 'PAVIMENTAÇÃO - 2026 - km (f)', 'PAVIMENTAÇÃO - 2026 - Ext. (km)', 
                    'PAVIMENTAÇÃO - 2026 - PERCENTUAL (%)', 'PAVIMENTAÇÃO - 2026 - (km)%', 'PAVIMENTAÇÃO - 2026 - FINANCEIRO (R$)', 
                    'PAVIMENTAÇÃO - Pós 2026 - Descrição', 'PAVIMENTAÇÃO - Pós 2026 - km (i)', 'PAVIMENTAÇÃO - Pós 2026 - km (f)', 
                    'PAVIMENTAÇÃO - Pós 2026 - Ext. (km)', 'PAVIMENTAÇÃO - Pós 2026 - PERCENTUAL (%)', 'PAVIMENTAÇÃO - Pós 2026 - (km)%', 
                    'PAVIMENTAÇÃO - Pós 2026 - FINANCEIRO (R$)', 'PAVIMENTAÇÃO - Pós 2026 - REL. FÍSICO (km)', 
                    'PAVIMENTAÇÃO - Pós 2026 - REL.FINANCEIRO R$)', 
                    'DUPLICAÇÃO - < 2022 - Descrição', 'DUPLICAÇÃO - < 2022 - km (i)', 'DUPLICAÇÃO - < 2022 - km (f)', 
                    'DUPLICAÇÃO - < 2022 - Ext. (km)', 'DUPLICAÇÃO - < 2022 - PERCENTUAL (%)', 'DUPLICAÇÃO - < 2022 - (km)%', 
                    'DUPLICAÇÃO - < 2022 - FINANCEIRO (R$)', 'DUPLICAÇÃO - 2023 - Descrição', 'DUPLICAÇÃO - 2023 - km (i)', 
                    'DUPLICAÇÃO - 2023 - km (f)', 'DUPLICAÇÃO - 2023 - Ext. (km)', 'DUPLICAÇÃO - 2023 - PERCENTUAL (%)', 
                    'DUPLICAÇÃO - 2023 - (km)%', 'DUPLICAÇÃO - 2023 - FINANCEIRO (R$)', 'DUPLICAÇÃO - 2024 - Descrição', 
                    'DUPLICAÇÃO - 2024 - km (i)', 'DUPLICAÇÃO - 2024 - km (f)', 'DUPLICAÇÃO - 2024 - Ext. (km)', 
                    'DUPLICAÇÃO - 2024 - PERCENTUAL (%)', 'DUPLICAÇÃO - 2024 - (km)%', 'DUPLICAÇÃO - 2024 - FINANCEIRO (R$)', 
                    'DUPLICAÇÃO - 2025 - Descrição', 'DUPLICAÇÃO - 2025 - km (i)', 'DUPLICAÇÃO - 2025 - km (f)', 
                    'DUPLICAÇÃO - 2025 - Ext. (km)', 'DUPLICAÇÃO - 2025 - PERCENTUAL (%)', 'DUPLICAÇÃO - 2025 - (km)%', 
                    'DUPLICAÇÃO - 2025 - FINANCEIRO (R$)', 'DUPLICAÇÃO - 2026 - Descrição', 'DUPLICAÇÃO - 2026 - km (i)', 
                    'DUPLICAÇÃO - 2026 - km (f)', 'DUPLICAÇÃO - 2026 - Ext. (km)', 'DUPLICAÇÃO - 2026 - PERCENTUAL (%)', 
                    'DUPLICAÇÃO - 2026 - (km)%', 'DUPLICAÇÃO - 2026 - FINANCEIRO (R$)', 'DUPLICAÇÃO - Pós 2026 - Descrição', 
                    'DUPLICAÇÃO - Pós 2026 - km (i)', 'DUPLICAÇÃO - Pós 2026 - km (f)', 'DUPLICAÇÃO - Pós 2026 - Ext. (km)', 
                    'DUPLICAÇÃO - Pós 2026 - PERCENTUAL (%)', 'DUPLICAÇÃO - Pós 2026 - (km)%', 'DUPLICAÇÃO - Pós 2026 - FINANCEIRO (R$)', 
                    'DUPLICAÇÃO - Pós 2026 - REL. FÍSICO (km)', 'DUPLICAÇÃO - Pós 2026 - REL.FINANCEIRO R$)', 
                    'OAE - < 2022 - Descrição', 
                    'OAE - < 2022 - km (i)', 'OAE - < 2022 - km (f)', 'OAE - < 2022 - Ext. (km)', 'OAE - < 2022 - PERCENTUAL (%)', 
                    'OAE - < 2022 - (km)%', 'OAE - < 2022 - FINANCEIRO (R$)', 'OAE - 2023 - Descrição', 'OAE - 2023 - km (i)',
                    'OAE - 2023 - km (f)', 'OAE - 2023 - Ext. (km)', 'OAE - 2023 - PERCENTUAL (%)', 'OAE - 2023 - (km)%', 
                    'OAE - 2023 - FINANCEIRO (R$)', 'OAE - 2024 - Descrição', 'OAE - 2024 - km (i)', 'OAE - 2024 - km (f)', 
                    'OAE - 2024 - Ext. (km)', 'OAE - 2024 - PERCENTUAL (%)', 'OAE - 2024 - (km)%', 'OAE - 2024 - FINANCEIRO (R$)', 
                    'OAE - 2025 - Descrição', 'OAE - 2025 - km (i)', 'OAE - 2025 - km (f)', 'OAE - 2025 - Ext. (km)', 
                    'OAE - 2025 - PERCENTUAL (%)', 'OAE - 2025 - (km)%', 'OAE - 2025 - FINANCEIRO (R$)', 'OAE - 2026 - Descrição', 
                    'OAE - 2026 - km (i)', 'OAE - 2026 - km (f)', 'OAE - 2026 - Ext. (km)', 'OAE - 2026 - PERCENTUAL (%)', 
                    'OAE - 2026 - (km)%', 'OAE - 2026 - FINANCEIRO (R$)', 'OAE - Pós 2026 - Descrição', 'OAE - Pós 2026 - km (i)', 
                    'OAE - Pós 2026 - km (f)', 'OAE - Pós 2026 - Ext. (km)', 'OAE - Pós 2026 - PERCENTUAL (%)', 'OAE - Pós 2026 - (km)%', 
                    'OAE - Pós 2026 - FINANCEIRO (R$)', 'OAE - Pós 2026 - REL. FÍSICO (km)', 'OAE - Pós 2026 - REL.FINANCEIRO R$)', 
                    'CONTORNO - < 2022 - Descrição', 'CONTORNO - < 2022 - km (i)', 'CONTORNO - < 2022 - km (f)', 
                    'CONTORNO - < 2022 - Ext. (km)', 'CONTORNO - < 2022 - PERCENTUAL (%)', 'CONTORNO - < 2022 - (km)%', 
                    'CONTORNO - < 2022 - FINANCEIRO (R$)', 'CONTORNO - 2023 - Descrição', 'CONTORNO - 2023 - km (i)', 
                    'CONTORNO - 2023 - km (f)', 'CONTORNO - 2023 - Ext. (km)', 'CONTORNO - 2023 - PERCENTUAL (%)', 
                    'CONTORNO - 2023 - (km)%', 'CONTORNO - 2023 - FINANCEIRO (R$)', 'CONTORNO - 2024 - Descrição', 
                    'CONTORNO - 2024 - km (i)', 'CONTORNO - 2024 - km (f)', 'CONTORNO - 2024 - Ext. (km)', 
                    'CONTORNO - 2024 - PERCENTUAL (%)', 'CONTORNO - 2024 - (km)%', 'CONTORNO - 2024 - FINANCEIRO (R$)', 
                    'CONTORNO - 2025 - Descrição', 'CONTORNO - 2025 - km (i)', 'CONTORNO - 2025 - km (f)', 
                    'CONTORNO - 2025 - Ext. (km)', 'CONTORNO - 2025 - PERCENTUAL (%)', 'CONTORNO - 2025 - (km)%', 
                    'CONTORNO - 2025 - FINANCEIRO (R$)', 'CONTORNO - 2026 - Descrição', 'CONTORNO - 2026 - km (i)', 
                    'CONTORNO - 2026 - km (f)', 'CONTORNO - 2026 - Ext. (km)', 'CONTORNO - 2026 - PERCENTUAL (%)', 
                    'CONTORNO - 2026 - (km)%', 'CONTORNO - 2026 - FINANCEIRO (R$)', 'CONTORNO - Pós 2026 - Descrição', 
                    'CONTORNO - Pós 2026 - km (i)', 'CONTORNO - Pós 2026 - km (f)', 'CONTORNO - Pós 2026 - Ext. (km)', 
                    'CONTORNO - Pós 2026 - PERCENTUAL (%)', 'CONTORNO - Pós 2026 - (km)%', 'CONTORNO - Pós 2026 - FINANCEIRO (R$)', 
                    'CONTORNO - Pós 2026 - REL. FÍSICO (km)', 'CONTORNO - Pós 2026 - REL.FINANCEIRO R$)', 
                    'FX ADICIONAL - < 2022 - Descrição', 'FX ADICIONAL - < 2022 - km (i)', 
                    'FX ADICIONAL - < 2022 - km (f)', 'FX ADICIONAL - < 2022 - Ext. (km)', 
                    'FX ADICIONAL - < 2022 - PERCENTUAL (%)', 'FX ADICIONAL - < 2022 - (km)%', 
                    'FX ADICIONAL - < 2022 - FINANCEIRO (R$)', 'FX ADICIONAL - 2023 - Descrição', 
                    'FX ADICIONAL - 2023 - km (i)', 'FX ADICIONAL - 2023 - km (f)', 'FX ADICIONAL - 2023 - Ext. (km)',
                    'FX ADICIONAL - 2023 - PERCENTUAL (%)', 'FX ADICIONAL - 2023 - (km)%', 'FX ADICIONAL - 2023 - FINANCEIRO (R$)', 
                    'FX ADICIONAL - 2024 - Descrição', 'FX ADICIONAL - 2024 - km (i)', 'FX ADICIONAL - 2024 - km (f)', 
                    'FX ADICIONAL - 2024 - Ext. (km)', 'FX ADICIONAL - 2024 - PERCENTUAL (%)', 'FX ADICIONAL - 2024 - (km)%', 
                    'FX ADICIONAL - 2024 - FINANCEIRO (R$)', 'FX ADICIONAL - 2025 - Descrição', 'FX ADICIONAL - 2025 - km (i)', 
                    'FX ADICIONAL - 2025 - km (f)', 'FX ADICIONAL - 2025 - Ext. (km)', 'FX ADICIONAL - 2025 - PERCENTUAL (%)', 
                    'FX ADICIONAL - 2025 - (km)%', 'FX ADICIONAL - 2025 - FINANCEIRO (R$)', 'FX ADICIONAL - 2026 - Descrição',
                    'FX ADICIONAL - 2026 - km (i)', 'FX ADICIONAL - 2026 - km (f)', 'FX ADICIONAL - 2026 - Ext. (km)', 
                    'FX ADICIONAL - 2026 - PERCENTUAL (%)', 'FX ADICIONAL - 2026 - (km)%', 'FX ADICIONAL - 2026 - FINANCEIRO (R$)', 
                    'FX ADICIONAL - Pós 2026 - Descrição', 'FX ADICIONAL - Pós 2026 - km (i)', 'FX ADICIONAL - Pós 2026 - km (f)', 
                    'FX ADICIONAL - Pós 2026 - Ext. (km)', 'FX ADICIONAL - Pós 2026 - PERCENTUAL (%)', 'FX ADICIONAL - Pós 2026 - (km)%', 
                    'FX ADICIONAL - Pós 2026 - FINANCEIRO (R$)', 'FX ADICIONAL - Pós 2026 - REL. FÍSICO (km)', 
                    'FX ADICIONAL - Pós 2026 - REL.FINANCEIRO R$)', 
                    'TERCEIRA FAIXA - < 2022 - Descrição', 
                    'TERCEIRA FAIXA - < 2022 - km (i)', 'TERCEIRA FAIXA - < 2022 - km (f)', 
                    'TERCEIRA FAIXA - < 2022 - Ext. (km)', 'TERCEIRA FAIXA - < 2022 - PERCENTUAL (%)',
                    'TERCEIRA FAIXA - < 2022 - (km)%', 'TERCEIRA FAIXA - < 2022 - FINANCEIRO (R$)', 
                    'TERCEIRA FAIXA - 2023 - Descrição', 'TERCEIRA FAIXA - 2023 - km (i)', 
                    'TERCEIRA FAIXA - 2023 - km (f)', 'TERCEIRA FAIXA - 2023 - Ext. (km)', 
                    'TERCEIRA FAIXA - 2023 - PERCENTUAL (%)', 'TERCEIRA FAIXA - 2023 - (km)%', 
                    'TERCEIRA FAIXA - 2023 - FINANCEIRO (R$)', 'TERCEIRA FAIXA - 2024 - Descrição', 
                    'TERCEIRA FAIXA - 2024 - km (i)', 'TERCEIRA FAIXA - 2024 - km (f)', 
                    'TERCEIRA FAIXA - 2024 - Ext. (km)', 'TERCEIRA FAIXA - 2024 - PERCENTUAL (%)', 
                    'TERCEIRA FAIXA - 2024 - (km)%', 'TERCEIRA FAIXA - 2024 - FINANCEIRO (R$)', 'TERCEIRA FAIXA - 2025 - Descrição', 
                    'TERCEIRA FAIXA - 2025 - km (i)', 'TERCEIRA FAIXA - 2025 - km (f)', 'TERCEIRA FAIXA - 2025 - Ext. (km)', 
                    'TERCEIRA FAIXA - 2025 - PERCENTUAL (%)', 'TERCEIRA FAIXA - 2025 - (km)%', 'TERCEIRA FAIXA - 2025 - FINANCEIRO (R$)', 
                    'TERCEIRA FAIXA - 2026 - Descrição', 'TERCEIRA FAIXA - 2026 - km (i)', 'TERCEIRA FAIXA - 2026 - km (f)', 
                    'TERCEIRA FAIXA - 2026 - Ext. (km)', 'TERCEIRA FAIXA - 2026 - PERCENTUAL (%)', 'TERCEIRA FAIXA - 2026 - (km)%', 
                    'TERCEIRA FAIXA - 2026 - FINANCEIRO (R$)', 'TERCEIRA FAIXA - Pós 2026 - Descrição', 'TERCEIRA FAIXA - Pós 2026 - km (i)', 
                    'TERCEIRA FAIXA - Pós 2026 - km (f)', 'TERCEIRA FAIXA - Pós 2026 - Ext. (km)', 
                    'TERCEIRA FAIXA - Pós 2026 - PERCENTUAL (%)', 'TERCEIRA FAIXA - Pós 2026 - (km)%', 
                    'TERCEIRA FAIXA - Pós 2026 - FINANCEIRO (R$)', 'TERCEIRA FAIXA - Pós 2026 - REL. FÍSICO (km)', 
                    'TERCEIRA FAIXA - Pós 2026 - REL.FINANCEIRO R$)']

# %%
# Criar um dicionário que mapeia 'Unnamed: x' para os novos nomes
rename_dict = {'Unnamed: {}'.format(i): name for i, name in enumerate(new_column_names, start=9)}

# Usar o método rename para renomear as colunas
df_per_copy.rename(columns=rename_dict, inplace=True)

# %%
df_per_copy['ESTRUTURADOR DO PROJETO'].unique()

# %%
df_per_copy['PROPONENTE'].unique()

# %%
df_per_copy['ESTRUTURADOR DO PROJETO'] = 'EPL'

# %%
df_per_copy['PROPONENTE'] = 'ANTT'

# %%
df_per_copy.index = range(1, len(df_per_copy) + 1)
df_per_copy.index.name = 'ID-ÚNICO'
df_per_copy = df_per_copy.reset_index()

# %%
df_per_copy

# %%
df_per_raw = df_per_copy

# %%
# df_per_copy.to_excel('Dados Gerados/RAW_PER (1).xlsx', index=False)

# %%
colunas_manter = [
    'ID-ÚNICO', 'SETOR', 'UF', 'ESTADO/LOTE', 'BR', 'EMPREENDIMENTO', 'PROPONENTE',
    'EXECUTOR (Grupo Controlador)', 'ESTRUTURADOR DO PROJETO'
]

# Use 'difference' para obter as colunas a serem unpivotadas
colunas_unpivot = df_per_copy.columns.difference(colunas_manter)

# Realizar o unpivot (melt) das outras colunas
df_per_copy = df_per_copy.melt(id_vars=colunas_manter, value_vars=colunas_unpivot,
                    var_name='Atributo', value_name='Valor')

# %%
df_per_copy

# %%
df_per_copy = df_per_copy.loc[df_per_copy['Valor'] != 0]
df_per_copy = df_per_copy.dropna(subset=['Valor'])

# %%
df_per_copy[['Atributo.1', 'Atributo.3', 'Atributo.2']] = df_per_copy['Atributo'].str.split(' - ', expand=True)


# %%
df_per_copy = df_per_copy[['ID-ÚNICO', 'SETOR', 'UF', 'ESTADO/LOTE', 'BR', 'EMPREENDIMENTO', 'PROPONENTE',
    'EXECUTOR (Grupo Controlador)', 'ESTRUTURADOR DO PROJETO', 'Atributo.1', 'Atributo.2', 'Atributo.3', 'Valor']]

# %%
'''# Converter colunas para 'category'
df_per_copy['SETOR'] =  df_per_copy['SETOR'].astype('category')
df_per_copy['UF'] =  df_per_copy['UF'].astype('category')
df_per_copy['PROPONENTE'] =  df_per_copy['PROPONENTE'].astype('category')
df_per_copy['EXECUTOR (Grupo Controlador)'] =  df_per_copy['EXECUTOR (Grupo Controlador)'].astype('category')
df_per_copy['ESTRUTURADOR DO PROJETO'] =  df_per_copy['ESTRUTURADOR DO PROJETO'].astype('category')
df_per_copy['Atributo.1'] =  df_per_copy['Atributo.1'].astype('category')
df_per_copy['Atributo.2'] =  df_per_copy['Atributo.2'].astype('category')

# Converter colunas para 'string'
df_per_copy['EMPREENDIMENTO'] =  df_per_copy['EMPREENDIMENTO'].astype('string') 
'''

# %%
df_per_copy

# %%
df_per_copy.to_excel('Dados Gerados/PER (1).xlsx', index=False)

# %% [markdown]
# ---

# %%
df_per_raw

# %%
# Drop na coluna 'Região' que não é usada.
df_per_raw.drop(columns=['ID-ÚNICO'], inplace=True)

# %%
df_per_raw.index = range(1, len(df_per_raw) + 1)
df_per_raw.index.name = 'ID-ÚNICO'
df_per_raw = df_per_raw.reset_index()

# %%
df_per_raw = df_per_raw.sort_values(by='ID-ÚNICO', ascending=True)

# %%
# Lista de strings que devem estar no nome das colunas para serem mantidas
keywords_to_keep = [
    'ID-ÚNICO', 'SETOR', 'EMPREENDIMENTO', 
    'Descrição', 'Ext. (km)', 'FINANCEIRO (R$)', 
    '(km)%'
]

# Identifica colunas que possuem qualquer uma das strings na lista de keywords
columns_to_keep = [col for col in df_per_raw.columns if any(keyword in col for keyword in keywords_to_keep)]

# %%
# Mantém apenas as colunas identificadas
df_per_raw = df_per_raw[columns_to_keep]

# %%
df_per_raw.columns

# %%
# Seleciona colunas da posição em diante
cols_a_verificar = df_per_raw.iloc[:, 3:]

# Remove linhas onde todas essas colunas estão vazias
df_per_raw = df_per_raw.dropna(subset=cols_a_verificar.columns, how='all')

# %%
df_per_raw1 = df_per_raw[['ID-ÚNICO', 'SETOR', 'EMPREENDIMENTO']]

# %%
colunas_manter = ['ID-ÚNICO', 'SETOR', 'EMPREENDIMENTO']

# Use 'difference' para obter as colunas a serem unpivotadas
colunas_unpivot = df_per_raw.columns.difference(colunas_manter)

# Realizar o unpivot (melt) das outras colunas
df_per_raw = df_per_raw.melt(id_vars=colunas_manter, value_vars=colunas_unpivot,
                    var_name='Atributo', value_name='Valor')

# %%
df_per_raw['SETOR'] = 'Rodoviário'

# %%
df_per_raw

# %%
df_per_raw[['Atributo.1', 'Atributo.3', 'Atributo.2']] = df_per_raw['Atributo'].str.split(' - ', expand=True)

# %%
df_per_raw = df_per_raw[['ID-ÚNICO', 'SETOR', 'EMPREENDIMENTO', 'Atributo.1', 'Atributo.2', 'Valor']]

# %%
df_per_raw.columns

# %%
df_per_raw

# %%
df_per_raw.dropna(subset=['Valor'])

# %%
# Pivotar o DataFrame
df_per_raw = df_per_raw.pivot_table(index=['ID-ÚNICO', 'Atributo.1'], 
                                            columns='Atributo.2', values='Valor', aggfunc='first').reset_index()

# %%
df_per_raw.columns

# %%
df_per_raw

# %%
# Remove a coluna de índice indesejada, se existir
if 'Atributo.2' in df_per_raw.columns:
    df_per_raw = df_per_raw.drop(columns='Atributo.2')

# %%
df_per_raw1

# %%
# Faz o merge com base na coluna 'ID-ÚNICO'
df_per_raw = pd.merge(df_per_raw1, df_per_raw, on='ID-ÚNICO', how='inner')

# %%
df_per_raw

# %%
df_per_raw = df_per_raw.dropna(subset=['Atributo.1'])

# %%
df_per_raw

# %%
df_per_raw.columns

# %%
df_per_raw = df_per_raw[['ID-ÚNICO', 'SETOR', 'EMPREENDIMENTO', 'Atributo.1', 'Descrição',
            'Ext. (km)', '(km)%', 'FINANCEIRO (R$)']]

# %%
# Renomear Colunas
df_per_raw.rename(columns={'ID-ÚNICO': 'ID-ÚNICO1'}, inplace=True)
df_per_raw.rename(columns={'SETOR': 'SETOR2'}, inplace=True)
df_per_raw.rename(columns={'EMPREENDIMENTO': 'EMPREENDIMENTO2'}, inplace=True)
df_per_raw.rename(columns={'Descrição': 'Descrição2'}, inplace=True)
df_per_raw.rename(columns={'Ext. (km)': 'Ext.(km)2'}, inplace=True)
df_per_raw.rename(columns={'FINANCEIRO (R$)': 'FINANCEIRO(R$)2'}, inplace=True)
df_per_raw.rename(columns={'(km)%': '(km)%2'}, inplace=True)

# %%
df_per_raw

# %%
df_per_raw.to_excel('PER(2).xlsx', sheet_name= 'PER(2)', index= False)


