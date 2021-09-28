import pygame
pygame.init()


W = 1000
H = 683
posX = 50
posY = 50
velX = 1
velY = 1
#obj = pygame.image.load('x.png')
#fundo = pygame.image.load('fundo.png')

janela = pygame.display.set_mode((W, H))
pygame.display.set_caption("Junim")

janela_aberta = True

while janela_aberta :
  pygame.time.delay(10)

  for event in pygame.event.get():
    if event.type == pygame.QUIT:
      janela_aberta = False



  if posX + velX + 100 > W or posX + velX < 0:
    velX = -velX
  if posY + velY + 100 > H or posY + velY < 0:
    velY = -velY

  posX = posX + velX
  posY = posY + velY
  
#   comandos = pygame.key.get_pressed()
#  if comandos[pygame.K_UP]:
#    posY -= velY
#  if comandos[pygame.K_DOWN]:
#    posY += velY
#  if comandos[pygame.K_LEFT]:
#    posX -= velX
#  if comandos[pygame.K_RIGHT]:
#    posX += velX      

  janela.fill(("#242937"))
  #janela.blit(fundo, (0,0))
  pygame.draw.rect(janela,"#DF2935", (posX, posY, 100, 100))
  pygame.draw.rect(janela,"#FA790F", (posX/200, posY+50, 100, 100))
  pygame.draw.rect(janela,"#F8D112", (2*posX, posY, 100, 100))
  pygame.draw.rect(janela,"#0CCE6B", (posX, 3*posY, 100, 100))
  pygame.draw.rect(janela,"#307CD7", (2*posX+1, 2*posY+1, 100, 100))
  
  #janela.blit(pygame.transform.rotate())
  pygame.display.update()

pygame.quit()
