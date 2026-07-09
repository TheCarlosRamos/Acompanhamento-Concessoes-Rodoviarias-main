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
# <h1><b><i>'À EXECUTAR (2026 DIANTE)'</b></i></h1>

# %%
#CHECKPOINT
df_a_executar_copy = df_a_executar.copy()

# %%
df_a_executar_copy = pd.read_excel(file_path, sheet_name='À EXECUTAR (2026 DIANTE)', header=2)

# %%
df_a_executar_copy

# %%
df_a_executar_copy = df_a_executar_copy.drop([0, 1, 2], axis = 0)

# %%
# Renomear Coluna ['EXECUTOR...']
df_a_executar_copy.rename(columns={'EXECUTOR             (Grupo Controlador)': 'EXECUTOR (Grupo Controlador)'}, inplace=True)

df_a_executar_copy.rename(columns={'PAVIMENTAÇÃO': 'PAVIMENTAÇÃO - 2026 - Descrição'}, inplace=True)
df_a_executar_copy.rename(columns={'DUPLICAÇÃO': 'DUPLICAÇÃO - 2026 - Descrição'}, inplace=True)
df_a_executar_copy.rename(columns={'OAE': 'OAE - 2026 - Descrição'}, inplace=True)
df_a_executar_copy.rename(columns={'CONTORNO': 'CONTORNO - 2026 - Descrição'}, inplace=True)
df_a_executar_copy.rename(columns={'FX ADICIONAL': 'FX ADICIONAL - 2026 - Descrição'}, inplace=True)
df_a_executar_copy.rename(columns={'TERCEIRA FAIXA': 'TERCEIRA FAIXA - 2026 - Descrição'}, inplace=True)

# %%
new_column_names = ['PAVIMENTAÇÃO - 2026 - km (i)', 
                    'PAVIMENTAÇÃO - 2026 - km (f)', 
                    'PAVIMENTAÇÃO - 2026 - Ext. (km)', 
                    'PAVIMENTAÇÃO - 2026 - PERCENTUAL (%)', 
                    'PAVIMENTAÇÃO - 2026 - (km)%', 
                    'PAVIMENTAÇÃO - 2026 - FINANCEIRO (R$)',
                    'PAVIMENTAÇÃO - 2026 - Descrição', 
                    'PAVIMENTAÇÃO - Pós 2026 - km (i)', 
                    'PAVIMENTAÇÃO - Pós 2026 - km (f)', 
                    'PAVIMENTAÇÃO - Pós 2026 - Ext. (km)', 
                    'PAVIMENTAÇÃO - Pós 2026 - PERCENTUAL (%)', 
                    'PAVIMENTAÇÃO - Pós 2026 - (km)%', 
                    'PAVIMENTAÇÃO - Pós 2026 - FINANCEIRO (R$)', 
                    'PAVIMENTAÇÃO - Pós 2026 - REL. FÍSICO (km)', 
                    'PAVIMENTAÇÃO - Pós 2026 - REL.FINANCEIRO (R$)',
                    'DUPLICAÇÃO - 2026 - Descrição', 
                    'DUPLICAÇÃO - 2026 - km (i)', 'DUPLICAÇÃO - 2026 - km (f)', 'DUPLICAÇÃO - 2026 - Ext. (km)', 
                    'DUPLICAÇÃO - 2026 - PERCENTUAL (%)', 'DUPLICAÇÃO - 2026 - (km)%', 'DUPLICAÇÃO - 2026 - FINANCEIRO (R$)', 
                    'DUPLICAÇÃO - Pós 2026 - Descrição', 'DUPLICAÇÃO - Pós 2026 - km (i)', 'DUPLICAÇÃO - Pós 2026 - km (f)', 
                    'DUPLICAÇÃO - Pós 2026 - Ext. (km)', 'DUPLICAÇÃO - Pós 2026 - PERCENTUAL (%)', 'DUPLICAÇÃO - Pós 2026 - (km)%', 
                    'DUPLICAÇÃO - Pós 2026 - FINANCEIRO (R$)', 'DUPLICAÇÃO - Pós 2026 - REL. FÍSICO (km)', 'DUPLICAÇÃO - Pós 2026 - REL.FINANCEIRO (R$)',
                    'OAE - 2026 - Descrição', 
                    'OAE - 2026 - km (i)', 'OAE - 2026 - km (f)', 'OAE - 2026 - Ext. (km)', 
                    'OAE - 2026 - PERCENTUAL (%)', 'OAE - 2026 - (km)%', 'OAE - 2026 - FINANCEIRO (R$)', 
                    'OAE - Pós 2026 - Descrição', 'OAE - Pós 2026 - km (i)', 'OAE - Pós 2026 - km (f)', 
                    'OAE - Pós 2026 - Ext. (km)', 'OAE - Pós 2026 - PERCENTUAL (%)', 'OAE - Pós 2026 - (km)%', 
                    'OAE - Pós 2026 - FINANCEIRO (R$)', 'OAE - Pós 2026 - REL. FÍSICO (km)', 'OAE - Pós 2026 - REL.FINANCEIRO (R$)',
                    'CONTORNO - 2026 - Descrição', 
                    'CONTORNO - 2026 - km (i)', 'CONTORNO - 2026 - km (f)', 'CONTORNO - 2026 - Ext. (km)', 
                    'CONTORNO - 2026 - PERCENTUAL (%)', 'CONTORNO - 2026 - (km)%', 'CONTORNO - 2026 - FINANCEIRO (R$)', 
                    'CONTORNO - Pós 2026 - Descrição', 'CONTORNO - Pós 2026 - km (i)', 'CONTORNO - Pós 2026 - km (f)', 
                    'CONTORNO - Pós 2026 - Ext. (km)', 'CONTORNO - Pós 2026 - PERCENTUAL (%)', 'CONTORNO - Pós 2026 - (km)%', 
                    'CONTORNO - Pós 2026 - FINANCEIRO (R$)', 'CONTORNO - Pós 2026 - REL. FÍSICO (km)', 'CONTORNO - Pós 2026 - REL.FINANCEIRO (R$)',
                    'FX ADICIONAL - 2026 - Descrição', 
                    'FX ADICIONAL - 2026 - km (i)', 'FX ADICIONAL - 2026 - km (f)', 'FX ADICIONAL - 2026 - Ext. (km)', 
                    'FX ADICIONAL - 2026 - PERCENTUAL (%)', 'FX ADICIONAL - 2026 - (km)%', 'FX ADICIONAL - 2026 - FINANCEIRO (R$)', 
                    'FX ADICIONAL - Pós 2026 - Descrição', 'FX ADICIONAL - Pós 2026 - km (i)', 'FX ADICIONAL - Pós 2026 - km (f)', 
                    'FX ADICIONAL - Pós 2026 - Ext. (km)', 'FX ADICIONAL - Pós 2026 - PERCENTUAL (%)', 'FX ADICIONAL - Pós 2026 - (km)%', 
                    'FX ADICIONAL - Pós 2026 - FINANCEIRO (R$)', 'FX ADICIONAL - Pós 2026 - REL. FÍSICO (km)', 'FX ADICIONAL - Pós 2026 - REL.FINANCEIRO (R$)',
                    'TERCEIRA FAIXA - 2026 - Descrição', 
                    'TERCEIRA FAIXA - 2026 - km (i)', 'TERCEIRA FAIXA - 2026 - km (f)', 'TERCEIRA FAIXA - 2026 - Ext. (km)', 
                    'TERCEIRA FAIXA - 2026 - PERCENTUAL (%)', 'TERCEIRA FAIXA - 2026 - (km)%', 'TERCEIRA FAIXA - 2026 - FINANCEIRO (R$)', 
                    'TERCEIRA FAIXA - Pós 2026 - Descrição', 'TERCEIRA FAIXA - Pós 2026 - km (i)', 'TERCEIRA FAIXA - Pós 2026 - km (f)', 
                    'TERCEIRA FAIXA - Pós 2026 - Ext. (km)', 'TERCEIRA FAIXA - Pós 2026 - PERCENTUAL (%)', 'TERCEIRA FAIXA - Pós 2026 - (km)%', 
                    'TERCEIRA FAIXA - Pós 2026 - FINANCEIRO (R$)', 'TERCEIRA FAIXA - Pós 2026 - REL. FÍSICO (km)', 'TERCEIRA FAIXA - Pós 2026 - REL.FINANCEIRO (R$)']

# %%
# Criar um dicionário que mapeia 'Unnamed: x' para os novos nomes
rename_dict = {'Unnamed: {}'.format(i): name for i, name in enumerate(new_column_names, start=9)}

# Usar o método rename para renomear as colunas
df_a_executar_copy.rename(columns=rename_dict, inplace=True)

# %%
df_a_executar_copy

# %%
df_a_executar_copy.dropna(subset=['ESTADO/LOTE'])

# %%
columns_to_fill = ['EMPREENDIMENTO']
df_a_executar_copy[columns_to_fill] = df_a_executar_copy.groupby((df_a_executar_copy['BR'] == 1).cumsum())[columns_to_fill].transform(lambda x: x.ffill())

# %%
# Manter apenas os valores válidos de estado/lote e propagar para baixo
df_a_executar_copy['ESTADO/LOTE'] = df_a_executar_copy['ESTADO/LOTE'].apply(
    lambda x: x if str(x) in [
        'RIO GRANDE DO SUL', 'MINAS GERAIS', 'GOIÁS', 'SANTA CATARINA', 'TOCANTINS',
        'GOIAS', 'RIO DE JANEIRO', 'MATO GROSSO', 'PARÁ', 'RIO/SP', 'SÃO PAULO',
        'PARANÁ', 'BAHIA', 'ESPIRITO SANTO', 'MATO GROSSO DO SUL'
    ] else pd.NA
)

# Forward fill (compatível com sua versão do pandas)
df_a_executar_copy['ESTADO/LOTE'] = df_a_executar_copy['ESTADO/LOTE'].ffill()

# %%
df_a_executar_copy = df_a_executar_copy[~df_a_executar_copy.apply(lambda row: row.astype(str).str.contains('TOTAL|SOMA').any(), axis=1)]

# %%
df_a_executar_copy.columns

# %%
df_a_executar_copy.index = range(1, len(df_a_executar_copy) + 1)
df_a_executar_copy.index.name = 'ID-ÚNICO'
df_a_executar_copy = df_a_executar_copy.reset_index()

# %%
df_a_executar_raw = df_a_executar_copy

# %%
# df_a_executar_copy.to_excel('Dados Gerados/RAW_À EXECUTAR (2026 DIANTE).xlsx', index=False)

# %%
colunas_manter = ['ID-ÚNICO', 'SETOR', 'UF', 'ESTADO/LOTE', 'BR', 'EMPREENDIMENTO',
       'PROPONENTE', 'EXECUTOR (Grupo Controlador)', 'ESTRUTURADOR DO PROJETO']

# Use 'difference' para obter as colunas a serem unpivotadas
colunas_unpivot = df_a_executar_copy.columns.difference(colunas_manter)

# Realizar o unpivot (melt) das outras colunas
df_a_executar_copy = df_a_executar_copy.melt(id_vars=colunas_manter, value_vars=colunas_unpivot,
                    var_name='Atributo', value_name='Valor')

# %%
df_a_executar_copy['EMPREENDIMENTO'].unique()

# %%
df_a_executar_copy = df_a_executar_copy.loc[df_a_executar_copy['Valor'] != 0]
df_a_executar_copy = df_a_executar_copy.dropna(subset=['Valor'])

# %%
df_a_executar_copy[['Atributo.1', 'Atributo.3', 'Atributo.2']] = df_a_executar_copy['Atributo'].str.split(' - ', expand=True)

# %%
# DROP de ['Atributo'] e Realocar ['Valor'] para o final da Tabela.

df_a_executar_copy = df_a_executar_copy[['ID-ÚNICO', 'SETOR', 'UF', 'ESTADO/LOTE', 'BR', 'EMPREENDIMENTO',
       'PROPONENTE', 'EXECUTOR (Grupo Controlador)', 'ESTRUTURADOR DO PROJETO'
       , 'Atributo.1', 'Atributo.2', 'Atributo.3', 'Valor']]

# %%
df_a_executar_copy['Atributo.1'].unique()

# %%
'''

df_a_executar_copy['Atributo.1'] = df_a_executar_copy['Atributo.1'].apply(
    lambda x: x.replace('CONTORNO ', 'CONTORNO')
              .replace('DUPLICAÇÃO ', 'DUPLICAÇÃO')
              .replace('FX ADICIONAL ', 'FX ADICIONAL')
              .replace('OAE ', 'OAE')
              .replace('PAVIMENTAÇÃO ', 'PAVIMENTAÇÃO')
              .replace('TERCEIRA FAIXA ', 'TERCEIRA FAIXA')
)
'''

# %%
df_a_executar_copy['Atributo.2'].unique()

# %%
# Converter colunas para 'category'
df_a_executar_copy['SETOR'] = df_a_executar_copy['SETOR'].astype('category')
df_a_executar_copy['UF'] = df_a_executar_copy['UF'].astype('category')
df_a_executar_copy['PROPONENTE'] = df_a_executar_copy['PROPONENTE'].astype('category')
df_a_executar_copy['EXECUTOR (Grupo Controlador)'] = df_a_executar_copy['EXECUTOR (Grupo Controlador)'].astype('category')
df_a_executar_copy['ESTRUTURADOR DO PROJETO'] = df_a_executar_copy['ESTRUTURADOR DO PROJETO'].astype('category')
df_a_executar_copy['Atributo.1'] = df_a_executar_copy['Atributo.1'].astype('category')
df_a_executar_copy['Atributo.2'] = df_a_executar_copy['Atributo.2'].astype('category')

# Converter colunas para 'string'
df_a_executar_copy['EMPREENDIMENTO'] = df_a_executar_copy['EMPREENDIMENTO'].astype('string') 


# %%
df_a_executar_copy = df_a_executar_copy.sort_values(by='ID-ÚNICO', ascending=True)

# %%
df_a_executar_copy

# %%
df_a_executar_copy.to_excel('Dados Gerados/À EXECUTAR (2026 DIANTE).xlsx', index=False)


