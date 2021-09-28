import pandas as pd
import os, sys

#!BANCO DE DADOS
#Explicação: Carregando Arquivo com os dados
caminho = os.path.dirname(__file__)
xls = pd.ExcelFile(os.path.join(caminho, "Coletar Modelos de Paredes.xlsm"))

# 'extrai' ΔY (Distancia entre as Linhas Base e a Plotagem)
DeltaY = pd.read_excel(xls, sheet_name='Paredes')
DeltaY = DeltaY.loc[0, 'DeltaY']

#'extrai' Paredes
Paredes = pd.read_excel(xls, sheet_name='Paredes')
Paredes  = Paredes[['ID', 'Xi', 'Yi', 'Xf', 'Yf', 'Lenght', 'Layer','Id Texto', 'Texto','ID esq', 'Layer esq']].dropna(how='all', axis=0)

#'extrai' Esquadrias
Esquadrias = pd.read_excel(xls, sheet_name='Esq').dropna(how='all', axis=0)

# 'extrai' LinFiadas
LinFiadas = pd.read_excel(xls, sheet_name='Fiadas')
LinFiadas = LinFiadas[['ID', 'Xi', 'Yi', 'Xf', 'Yf', 'Tipo', 'Layer', 'NúmFiada', 'ID par', 'Check']]
LinFiadas = LinFiadas.dropna(how='all', axis =0)

# 'extrai' Blocos
Blocos = pd.read_excel(xls, sheet_name='Blocos')
Blocos = Blocos[['ID', 'Xi', 'Yi', 'Xf', 'Yf', 'Tipo', 'Layer', 'Lenght', 'ID fiada', 'ID par', 'Check']]
Blocos = Blocos.dropna(how='all', axis =0)

class DadosPar:
  def __init__(self, NomeParede):
    
    self.NomeParede = NomeParede

    self.DELTAY = DeltaY
    self.Par = Paredes
    self.Esq = Esquadrias
    self.Fiadas = LinFiadas
    self.Blocos = Blocos

  def locEsq(self):
    self.Pesq = True
    self.i = 0

    while self.Pesq:
      if self.Par.loc[i, 'Texto'] == self.NomeParede:
        self.ParSelecID = self.Par.loc[i, 'ID']
        self.ParSelecXi = self.Par.loc[i, 'Xi']
        self.ParSelecYi = self.Par.loc[i, 'Yi']
        self.Pesq = False
      else:
        self.i += 1

    #- Procurando ESQUADRIAS
    EsqSelec = self.Esq.loc[self.Esq['ID par'] == self.ParSelecID ]
    EsqSelec['Xid'] = round((self.Esq['Xi'] - self.ParSelecXi) * 100, 0)
    EsqSelec['Yid'] = round(self.Esq['Peitoril'], 0)

  def locFiadas(self):
    self.Pesq = True
    self.i = 0

    while self.Pesq:
      if self.Par.loc[i, 'Texto'] == self.NomeParede:
        self.ParSelecID = self.Par.loc[i, 'ID']
        self.ParSelecXi = self.Par.loc[i, 'Xi']
        self.ParSelecYi = self.Par.loc[i, 'Yi']
        self.Pesq = False
      else:
        self.i += 1

    #- Procurando FIADAS
    FiadaSelc = self.Fiadas.loc[self.Fiadas['ID par'] == self.ParSelecID]
    FiadaSelc['Xid'] = round((self.Fiadas['Xi'] - self.ParSelecXi) * 100, 0)
    FiadaSelc['Yid'] = round((self.Fiadas['Yi'] - self.ParSelecYi - DeltaY) * 100 - 1, 0)
  
  def locBlocos(self):
    self.Pesq = True
    self.i = 0

    while self.Pesq:
      if self.Par.loc[i, 'Texto'] == self.NomeParede:
        self.ParSelecID = self.Par.loc[i, 'ID']
        self.ParSelecXi = self.Par.loc[i, 'Xi']
        self.ParSelecYi = self.Par.loc[i, 'Yi']
        self.Pesq = False
      else:
        self.i += 1

    #- Procurando BLOCOS
    BlocosSelc = Blocos.loc[Blocos['ID par'] == self.ParSelecID]
    BlocosSelc['Xid'] = round((BlocosSelc['Xi'] - self.ParSelecXi) * 100, 0)
    BlocosSelc['Yid'] = round((BlocosSelc['Yi'] - self.ParSelecYi - DeltaY) * 100 - 1, 0)

PAR = DadosPar("PAREDE 4 (7x)")
ESQ = PAR.locEsq
print(ESQ.loc[0, 'ID'])
