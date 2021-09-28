import pygame
import os

#!CONTANTES
#?Dados IA
IA_JOGANDO = True
GERACAO = 0
#-Dimensões da Tela
TELA_LARG = 500
TELA_ALT = 800
#-Images
caminho = os.path.dirname(__file__)
IMAGEM_CANO = pygame.transform.scale2x(pygame.image.load(os.path.join(caminho, 'image', 'pipe.png')))
IMAGEM_CHAO = pygame.transform.scale2x(pygame.image.load(os.path.join(caminho, 'image', 'base.png')))
IMAGEM_BACKGROUND =  pygame.transform.scale2x(pygame.image.load(os.path.join(caminho, 'image', 'bg.png')))
IMAGEM_PASSARO = [
  pygame.transform.scale2x(pygame.image.load(os.path.join(caminho, 'image', 'bird1.png'))),
  pygame.transform.scale2x(pygame.image.load(os.path.join(caminho, 'image', 'bird2.png'))),
  pygame.transform.scale2x(pygame.image.load(os.path.join(caminho, 'image', 'bird3.png')))
] #*é uma lista para aparentar movimento no pássaro
#Fonte do Placar
pygame.font.init()
FONTE_PONTOS = pygame.font.SysFont('Calibri', 50)


#!OBJETOS
class Passaro:
  IMGS = IMAGEM_PASSARO
  #Animações da rotação
  ROTACAO_MAX = 25
  VELOCIDADE_ROTACAO = 20
  TEMPO_ANIMACAO = 5

  def __init__(self, x, y):
    self.x = x
    self.y = y
    self.angulo = 0
    self.velocidade = 0
    self.altura = self.y
    self.tempo = 0
    self.contagem_imagem = 0
    self.imagem = self.IMGS[0]

  def pular(self):
    self.velocidade = -10.5
    self.tempo = 0
    self.altura = self.y

  def  mover(self):
    #Calcular o deslocamento
    self.tempo += 1
    deslocamento = 1.5 * (self.tempo**2) + self.velocidade * self.tempo

    #restringir deslocamento
    if deslocamento > 16:
      deslocamento = 16
    elif deslocamento < 0:
      deslocamento -= 2

    self.y += deslocamento

    #o angulo do passaro
    if deslocamento < 0 or self.y < (self.altura + 50):
      if self.angulo < self.ROTACAO_MAX:
        self.angulo = self.ROTACAO_MAX
    else:
      if self.angulo > -90:
        self.angulo -= self.VELOCIDADE_ROTACAO

  def desenhar(self, tela):
    #definir qual imagem do passaro vai usar
    self.contagem_imagem += 1
    #Alterando a imagem selecionada
    if self.contagem_imagem < self.TEMPO_ANIMACAO:
        self.imagem = self.IMGS[0]
    elif self.contagem_imagem < self.TEMPO_ANIMACAO*2:
        self.imagem = self.IMGS[1]
    elif self.contagem_imagem < self.TEMPO_ANIMACAO*3:
        self.imagem = self.IMGS[2]
    elif self.contagem_imagem < self.TEMPO_ANIMACAO*4:
        self.imagem = self.IMGS[1]
    elif self.contagem_imagem >= self.TEMPO_ANIMACAO*4 + 1:
        self.imagem = self.IMGS[0]
        self.contagem_imagem = 0


    #Se o passaro tiver caindo eu não vou bater asa
    if self.angulo <= -80:
      self.imagem = self.IMGS[1]
      self.contagem_imagem =self.TEMPO_ANIMACAO*2


    #desenhar a imagem
    imagem_rotacionada = pygame.transform.rotate(self.imagem, self.angulo)
    pos_centro_imagem = self.imagem.get_rect(topleft = (self.x, self.y)).center
    retangulo = imagem_rotacionada.get_rect(center = pos_centro_imagem)
    tela.blit(imagem_rotacionada, retangulo.topleft)

  def get_mask(self):
    return pygame.mask.from_surface(self.imagem)


class Cano:
  DISTANCIA  = 200
  VELOCIDADE  = 5

  def __init__(self, x):
    self.x = x
    self.altura = 0
    self.pos_topo = 0
    self.pos_base = 0
    self.CANO_TOPO = pygame.transform.flip(IMAGEM_CANO, False, True)
    self.CANO_BASE = IMAGEM_CANO
    self.passou = False
    self.definir_altura()

  def definir_altura(self):
    self.altura = random.randrange(50, 450)
    self.pos_topo = self.altura - self.CANO_TOPO.get_height()
    self.pos_base = self.altura + self.DISTANCIA

  def mover(self):
    self.x -= self.VELOCIDADE

  def desenhar(self, tela):
    tela.blit(self.CANO_TOPO, (self.x, self.pos_topo))
    tela.blit(self.CANO_BASE, (self.x, self.pos_base))

  def colidir(self, passaro):
    passaro_mask = passaro.get_mask()
    topo_mask = pygame.mask.from_surface(self.CANO_TOPO)
    base_mask = pygame.mask.from_surface(self.CANO_BASE)

    distancia_topo = (self.x - passaro.x, self.pos_topo - round(passaro.y))
    distancia_base = (self.x - passaro.x, self.pos_base - round(passaro.y))
    #os valores a seguir são True/False
    topo_ponto = passaro_mask.overlap(topo_mask, distancia_topo)
    base_ponto = passaro_mask.overlap(topo_mask, distancia_base)

    if base_ponto or topo_ponto:
      return True
    else:
      return False
    

class Chao:
  VELOCIDADE = 5
  LARG = IMAGEM_CHAO.get_width()
  IMAGEM = IMAGEM_CHAO

  def __init__(self, y):
    self.y = y
    self.x1 = 0
    self.x2 = self.LARG

  def mover(self):
    self.x1 -= self.VELOCIDADE
    self.x2 -= self.VELOCIDADE

    #verificar se chao saiu da tela
    if self.x1 + self.LARG < 0:
      self.x1 = self.x2 + self.LARG
    if self.x2 + self.LARG < 0:
      self.x2 = self.x1 + self.LARG
    
  def desenhar(self, tela):
    tela.blit(self.IMAGEM, (self.x1, self.y))
    tela.blit(self.IMAGEM, (self.x2, self.y))


#DESENHAR TELA DO JOGO
def desenharTela(tela, passaros, canos, chao, pontos):
  tela.blit(IMAGEM_BACKGROUND, (0, 0))

  for passaro in passaros:
    passaro.desenhar(tela)
  for cano in canos:
    cano.desenhar(tela)

  texto = FONTE_PONTOS.render(f"Pontuação: {pontos}", 1, ("#FFFFFF"))
  tela.blit(texto, (TELA_LARG - 10 - texto.get_width(), 10))
  chao.desenhar(tela)

  #Atualizar tela
  pygame.display.update()


#Fazendo a bagaça rodar
def main():
  passaros = [Passaro(230, 350)]
  chao = Chao(730)
  canos = [Cano(700)]
  tela = pygame.display.set_mode((TELA_LARG, TELA_ALT))
  pontos = 0
  relogio = pygame.time.Clock()

  #Loop pra manter o jogo aberto
  rodando = True
  while rodando:
    relogio.tick(30)                    

    #Verificar cada evento = tecla precionada
    for evento in pygame.event.get():
      #Botão de fechar
      if evento.type == pygame.QUIT:
        rodando = False
        pygame.quit()
        quit()
      #Botão pra pular (ESPAÇO)
      if evento.type == pygame.KEYDOWN:
        if evento.key == pygame.K_SPACE:
          for passaro in passaros:
            passaro.pular()

    #Mover coisas
    for passaro in passaros:
      passaro.mover()
    
    chao.mover()

    add_cano = False
    remover_canos= []
    for cano in canos:
      for i, passaro in enumerate(passaros):
        #se passaro bateu
        if cano.colidir(passaro):
          passaros.pop(i)
        #verifica se o passaro passou
        if not cano.passou and  passaro.x > cano.x:
          cano.passou = True
          add_cano = True
      cano.mover()

      if cano.x + cano.CANO_TOPO.get_width() < 0:
        remover_canos.append(cano)
    
    if add_cano:
      pontos += 1
      canos.append(Cano(600))

    for cano in remover_canos:
      canos.remove(cano)

    
    for i, passaro in enumerate(passaros):
      if (passaro.y + passaro.imagem.get_height()) > chao.y or passaro.y < 0:
        passaros.pop(i)
        
    desenharTela(tela, passaros, canos, chao, pontos)

if __name__ == '__main__':
  main()
