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

# constant
timer_pop_character = 1
timer_pop_space = 5
timer_average_character = 2.2
timer_threshold_character = 1.0
timer_absolute_value = 0.5

# output format
output_start_format = '\033[1;30;46m'
output_end_format = '\033[0m'

# exception output
output_exception_format = '\n\nThere is a typo! Restart~\n'


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
            if max_space / min_space >= timer_average_character or timestamps >= timer_threshold_character:
                decoder()
    except NameError:
        print(output_exception_format)
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
    morse_timer = threading.Timer(timer_pop_character, decoder)
    morse_timer.start()
    morse_space_timer = threading.Timer(timer_pop_space, plugin_space)
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
    format_code = output_start_format + str(' ') + output_end_format
    print(format_code, end='')


# translate morse code into characters
def decoder():
    global character
    global space
    global input
    try:
        if len(input) > 0:
            input = []
        average_character = sum(character) / len(character)
        for value in character:
            if value / average_character >= timer_average_character or value >= timer_absolute_value:
                input.append('-')
            else:
                input.append('.')
        character = []
        space = []
        code = ''.join(input)
        format_code = output_start_format + str(mtalk.decode(code)) + output_end_format
        print(format_code, end='')
    except KeyError:
        print(output_exception_format)
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
