from pynput import keyboard
import time
import morse_talk as mtalk
import pygame

character = []
space = []
input = []

last_release_time = time.time()
last_press_time = time.time()

pygame.mixer.init()
channel = pygame.mixer.Channel(2)
sound = pygame.mixer.Sound('d:/tmp/663.wav')


def on_press(key):
    global last_release_time
    global last_press_time
    global character
    global space
    global input
    last_press_time = time.time()
    if not channel.get_busy():
        channel.play(sound)
    timestamps = last_press_time - last_release_time
    try:
        if len(character) >= 1:
            space.append(timestamps)
        if len(space) > 1:
            max_space = max(space)
            min_space = min(space)
            if max_space / min_space >= 2.2 or timestamps >= 1.0:
                # print('space' + str(space))
                # print('character' + str(character))
                if len(input) > 0:
                    input = []
                avg_charater = sum(character) / len(character)
                for value in character:
                    if value / avg_charater >= 2.2 or value >= 0.5:
                        input.append('-')
                    else:
                        input.append('.')
                character = []
                space = []
                code = ''.join(input)
                print(mtalk.decode(code), end='')
    except:
        print('Error')
    return False


def on_release(key):
    global last_release_time
    global last_press_time
    channel.stop()
    last_release_time = time.time()
    timestamps = last_release_time - last_press_time
    character.append(timestamps)
    return False


def main():
    while 1:
        with keyboard.Listener(
                on_press=on_press) as listener:
            listener.join()
        with keyboard.Listener(
                on_release=on_release) as listener:
            listener.join()


if __name__ == '__main__':
    main()
