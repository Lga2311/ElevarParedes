import pandas as pd
import os, sys

#!BANCO DE DADOS
#Explicação: Carregando Arquivo com os dados
caminho = os.path.dirname(__file__)

# 'extrai' ΔY (Distancia entre as Linhas Base e a Plotagem)
DeltaY = pd.read_csv(os.path.join(caminho, 'Dados CSV', 'paredes.csv'), sep=';')
DeltaY = DeltaY.loc[0, 'DeltaY']

# 'extrai' Paredes
Paredes_df = pd.read_csv(os.path.join(caminho, 'Dados CSV', 'paredes.csv'), sep=';')
Paredes_df  = Paredes_df.dropna(how='all', axis=0)
Paredes_df  = Paredes_df.dropna(how='all', axis=1)
Paredes_df  = Paredes_df.drop('DeltaY', axis=1)

# 'extrai' Esquadrias
Esquadrias_df = pd.read_csv(os.path.join(caminho, 'Dados CSV', 'esquadrias.csv'), sep=';')

# 'extrai' LinFiadas
Fiadas_df = pd.read_csv(os.path.join(caminho, 'Dados CSV', 'fiadas.csv'), sep=';')
Fiadas_df = Fiadas_df.dropna(how='all', axis=0)
Fiadas_df = Fiadas_df.dropna(how='all', axis=1)

# 'extrai' Blocos
Blocos_df = pd.read_csv(os.path.join(caminho, 'Dados CSV', 'blocos.csv'), sep=';')
Blocos_df = Blocos_df.dropna(how='all', axis=0)
Blocos_df = Blocos_df.dropna(how='all', axis=1)



#Explicação: Transformando Tabelas em DICIONÁRIOS
# Paredes
paredes = {}
for par in Paredes_df['ID par']:
  paredes[par] = Paredes_df.loc[Paredes_df['ID par'] == par]
  paredes[par] = paredes[par].reset_index()

# esquadrias
esquadrias = {}
for par in Paredes_df['ID par']:
  esquadrias[par] = Esquadrias_df.loc[Esquadrias_df['ID par'] == par]
  esquadrias[par] = esquadrias[par].reset_index()

# fiadas
fiadas = {}
for par in Fiadas_df['ID par']:
  fiadas[par] = Fiadas_df.loc[Fiadas_df['ID par'] == par]
  fiadas[par] = fiadas[par].reset_index()

# blocos
blocos = {}
for par in Blocos_df['ID par']:
  blocos[par] = Blocos_df.loc[Blocos_df['ID par'] == par]
  blocos[par] = blocos[par].reset_index()

#Explicação: Criando funções para carregar as tabelas/dicionarios
# Paredes
def carregar_paredes():
  return paredes

def id_pars():
  pars = []
  for i in carregar_paredes().keys():
    pars.append(i)
    
  return pars

# esquadrias
def carregar_esquadrias():
  return esquadrias

# fiadas
def carregar_fiadas():
  return fiadas

# blocos
def carregar_blocos():
  return blocos

