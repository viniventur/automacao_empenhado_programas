import os
import warnings
import pandas as pd
import requests
from bs4 import BeautifulSoup
import wget
import ssl
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

data_atual = datetime.now()
data_ontem = data_atual - timedelta(days=1)
data_ontem = data_ontem.date().strftime('%d-%m-%Y')

ssl._create_default_https_context = ssl._create_unverified_context

url = f'https://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DESPESAS/despesa_empenhado_liquidado_pago_2024_siafe_gerado_em_{data_ontem}.xlsx'
arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'despesa_empenhado_liquidado_pago_2024_siafe_gerado_em_{data_ontem}.xlsx')
arquivo_csv = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'despesa_empenhado_liquidado_pago_2024_siafe_gerado_em_{data_ontem}.csv')
base_despesas_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'base_despesas1.xlsx')

# Baixar o arquivo
wget.download(url, arquivo)
df = pd.read_excel(arquivo)


df_f = df[['DESCRICAO_UO', 'PT', 'PT_DESCRICAO', 'NATUREZA6', 'VALOR_EMPENHADO', 'MES']]
for i in ['DESCRICAO_UO', 'PT', 'PT_DESCRICAO', 'NATUREZA6', 'MES']:
    df_f[i] = df_f[i].astype(str)

# concatenando as colunas para o merge
df_x = df_f
df_x['concat'] = df_x['DESCRICAO_UO']
for i in ['PT', 'NATUREZA6']:
    df_x['concat'] = df_x['concat'] + df_x[i]

# base com as informacoes
info = pd.read_excel('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python/Base acompanhamento Layne.xlsx')

# fazendo o merge para cada valor e cada mês
#info[['1', '2', '3']] = ''
for i in df_f['MES'].unique():
    info[i] = ''
    teste = df_x.loc[df_x['MES'] == i]
    for j in info['concat'].values:
     info.loc[info.concat == j, i] = teste.loc[teste.concat == j]['VALOR_EMPENHADO'].sum()
info
