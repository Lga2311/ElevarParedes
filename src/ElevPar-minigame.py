import sys, os
sys.path.insert(0, 'DataBase')
from DataBase.DadosParedes import id_pars, carregar_paredes, carregar_esquadrias, carregar_fiadas, carregar_blocos
import pygame
import random

#* CONSTANTES
TELA_ALT = 1010
TELA_LARG = 1920
ESP_LINHA = 2
XI = 200
YI = TELA_ALT - (200 + 10 + 1)

#- Carregando Dados das Paredes
paredesBD = carregar_paredes()
esqBD = carregar_esquadrias()
fiadasBD = carregar_fiadas()
blocosBD = carregar_blocos()

pars = id_pars()
pars_quant = len(pars)
i_par_rand = random.randint(0, pars_quant-1)
parede_selecionada = id_pars()[i_par_rand]
paredesBD = paredesBD[parede_selecionada]
esqBD = esqBD[parede_selecionada]
fiadasBD = fiadasBD[parede_selecionada]
blocosBD = blocosBD[parede_selecionada]

class Gabarito: 

  def __init__(self, Comp, nFiadas):
    self.comp = Comp
    self.nFiadas = nFiadas
    self.xi = 200
    self.xf = self.xi + self.comp
    self.yi = TELA_ALT - (200-9)
    self.yf = self.yi - (40 + 20 * self.nFiadas)

  def Desenhar(self, Janela):
    pts = ((self.xi, self.yf), 
           (self.xi, self.yi), 
           (self.xf, self.yi), 
           (self.xf, self.yf))
    #//pygame.draw.polygon(Janela, "#F8D112", pts, ESP_LINHA)
    pygame.draw.lines(Janela, "#FF0000", False, pts, ESP_LINHA)

class Esq: 

  def __init__(self, xi, yi, Larg, Alt, Peitoril):
    self.Larg = Larg
    self.Alt = Alt
    self.Peitoril = Peitoril
    self.xi = xi +  200
    self.yi = TELA_ALT - (200-9) - yi
    self.xf = self.xi + self.Larg - 1
    self.yf = self.yi - self.Alt

  def Desenhar(self, Janela):
    pts = ((self.xi, self.yf), 
           (self.xi, self.yi),
           (self.xf, self.yi), 
           (self.xf, self.yf))
    #//pygame.draw.polygon(Janela, "#F8D112", pts, ESP_LINHA)
    pygame.draw.lines(Janela, "#FF0000", True, pts, ESP_LINHA)
    
class Blocos:

  def __init__(self, Tipo, x, y, nFiada, comp):
    self.Tipo = Tipo
    self.x = x
    self.y = y
    self.yi = TELA_ALT - (200 + 10 + 20 + y)
    self.nFiada = nFiada
    self.comp = comp
    #//self.Visível = True

  def Desenhar(self, Janela):
    self.cor = "#2A3041"
    self.alt = 19 
    #//self.dim = (self.xi, self.yi-(self.alt+1)*self.nFiada, self.comp, self.alt)
    self.dim = (self.x, self.y, self.comp, self.alt)
    
    #Definindo Cores dos Blocos
    if self.Tipo == "B39" or self.comp == 39:
      self.cor = "#FFFF00"
    elif self.Tipo == "B19" or self.comp == 19:
      self.cor = "#00FF00"
    elif self.Tipo == "B34" or self.comp == 34:
      self.cor = "#00FFFF"
    elif self.Tipo == "B09" or self.comp == 9:
      self.cor = "#FF6600"
    elif self.Tipo == "B04" or self.comp == 4:
      self.cor = "#FF00FF"
    elif self.Tipo == "B54" or self.comp == 54:
      self.cor = "#FF0000"
    elif self.Tipo == "B29" or self.comp == 29:
      self.cor = "#6600FF"
    elif self.Tipo == "BP" or self.comp == 14:
      self.cor = "#107C41"


    pygame.draw.rect(Janela, self.cor, self.dim)

  def DesenharResposta(self, Janela, Visivel):
    self.Visivel = Visivel
    self.cor = "#242937"
    self.alt = 19 
    self.comp = self.comp
    #//self.dim = (self.xi, self.yi-(self.alt+1)*self.nFiada, self.comp, self.alt)
    self.dim = (self.x, self.y, self.comp, self.alt)

    #Definindo Cores dos Blocos
    if self.Visivel:
      self.cor = "#2A3041"  #Cor alternativa: "#202531"
    else:
      self.cor = "#242937"

    pygame.draw.rect(Janela, self.cor, self.dim)
  
def desenharTela(Tela, BlocosResposta, EstadoResposta, Esquadrias, Gabarito):
  pygame.display.set_caption("Elevar Paredes - MiniGame")
  Tela.fill(("#242937"))

  #Explicação: Respostas da Parede selecionada
  for i in range(len(BlocosResposta)):
    tipo = BlocosResposta.loc[i, "Layer"]
    comp = BlocosResposta.loc[i, "Lenght"]
    fiada = BlocosResposta.loc[i, "YiDelta"]/20 #? Pra quê q isso serve msm???
    X = XI + BlocosResposta.loc[i, "XiDelta"]
    Y = YI - BlocosResposta.loc[i, "YiDelta"]
    
    Blocos(tipo, X, Y, fiada, comp).DesenharResposta(Tela, EstadoResposta)

  #Explicação: Desenhando Esquadrias
  i=0
  for i in range(len(Esquadrias)):
    xi = int(Esquadrias.loc[i, 'XiDelta'])
    yi = Esquadrias.loc[i, 'YiDelta']
    larg = Esquadrias.loc[i,'Largura']
    alt = Esquadrias.loc[i, 'Altura']
    peitoril = Esquadrias.loc[i, 'Peitoril']
    
    Esq(xi, yi, larg, alt, peitoril).Desenhar(Tela)
    

  Gabarito.Desenhar(Tela)
  

  pygame.display.update()


def main():
  #Config da Tela
  TELA = pygame.display.set_mode((TELA_LARG, TELA_ALT))
  VISIVEL, INVISIVEL = True, False
  ÉÍNICIO = True
  estado = VISIVEL
  nF = 1
  i_par = i_par_rand
  i_par_max = pars_quant - 1

  # Add Relogio
  timer = pygame.time.Clock()
  rodando = True

  while rodando:
    timer.tick(60)
	  
    comandos = pygame.key.get_pressed()   
    for e in pygame.event.get():
      if e.type == pygame.QUIT: rodando = False
      if e.type == pygame.KEYDOWN:
          if e.key == pygame.K_s: estado = INVISIVEL
          if e.key == pygame.K_d: estado = VISIVEL
          if e.key == pygame.K_p:
            i_par_rand_nova = random.randint(0, pars_quant-1)

            ÉÍNICIO = False
            paredesBD_nova = carregar_paredes()
            esqBD_nova = carregar_esquadrias()
            fiadasBD_nova = carregar_fiadas()
            blocosBD_nova = carregar_blocos()

            parede_selecionada = pars[i_par_rand_nova]

            paredesBD_nova = paredesBD_nova[parede_selecionada]
            esqBD_nova = esqBD_nova[parede_selecionada]
            fiadasBD_nova = fiadasBD_nova[parede_selecionada]
            blocosBD_nova = blocosBD_nova[parede_selecionada]

            nome = paredesBD_nova.loc[0, 'Texto']
            print(f'{i_par_rand_nova} - {parede_selecionada}: {nome}')
    #Jogador

        

    if ÉÍNICIO:
      #Elementos Add
      gabarito = Gabarito(paredesBD['Lenght']*100, fiadasBD['NúmFiada'].max())
      desenharTela(TELA, blocosBD, estado, esqBD, gabarito)
      
    else:
      #Elementos Add
      gabarito = Gabarito(paredesBD_nova['Lenght']*100, fiadasBD_nova['NúmFiada'].max())
      desenharTela(TELA, blocosBD_nova, estado, esqBD_nova, gabarito)




#iniciando aplicação
if __name__ == "__main__":
  main()

