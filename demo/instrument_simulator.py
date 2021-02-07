import pygame
pygame.mixer.init()
channel = pygame.mixer.Channel(2)
sound = pygame.mixer.Sound('d:/tmp/morse.wav')
while 1:
    if channel.get_busy() == False:
        channel.play(sound)