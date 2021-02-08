from pynput import keyboard
import time
import morse_talk as mtalk
import pygame
import threading

# List Container consists of intervals between press and release
character = []
# List Container consists of intervals between release and press
space = []
# List Container consists of . and - messages
input = []

# release time
last_release_time = time.time()
# press time
last_press_time = time.time()

# initialize  sound effect
pygame.mixer.init()
channel = pygame.mixer.Channel(2)
sound = pygame.mixer.Sound('../source/morse.wav')


# press listener
def on_press(key):
    global last_release_time
    global last_press_time
    global character
    global space
    global input
    try:
        morse_timer.cancel()
        morse_space_timer.cancel()
    except NameError:
        pass
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
                decoder()
    except NameError:
        print('\n\nThere is a typo! Restart~\n')
        character = []
        space = []
        input = []
    return False


# release listener
def on_release(key):
    global last_release_time
    global last_press_time
    global character
    global space
    global input
    global morse_timer
    global morse_space_timer
    channel.stop()
    last_release_time = time.time()
    timestamps = last_release_time - last_press_time
    character.append(timestamps)
    morse_timer = threading.Timer(1, decoder)
    morse_timer.start()
    morse_space_timer = threading.Timer(5, plugin_space)
    morse_space_timer.start()
    return False


# plug in space
def plugin_space():
    global character
    global space
    global input
    input = []
    character = []
    space = []
    format_code = '\033[1;30;46m' + str(' ') + '\033[0m'
    print(format_code, end='')


# translate morse code into character
def decoder():
    global character
    global space
    global input
    try:
        if len(input) > 0:
            input = []
        average_character = sum(character) / len(character)
        for value in character:
            if value / average_character >= 2.2 or value >= 0.5:
                input.append('-')
            else:
                input.append('.')
        character = []
        space = []
        code = ''.join(input)
        format_code = '\033[1;30;46m' + str(mtalk.decode(code)) + '\033[0m'
        print(format_code, end='')
    except KeyError:
        print('\n\nThere is a typo! Restart~\n')
        character = []
        space = []
        input = []


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
