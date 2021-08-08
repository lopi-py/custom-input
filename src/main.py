from win32api import GetKeyState, MessageBox
from win32console import SetConsoleTitle, GetConsoleWindow
from win32con import MB_OK, MB_ICONERROR
from termcolor import colored
from colorama import init as colorama_init
colorama_init()
from sys import stdout, stdin, exit
from os import path, system

import keyboard
import time
import json

CAPS_LOCK_VK = 0x14
REFRESH_RATE = 0.001 # enable and disable detection refresh rate
PRESETS_FILE = "presets.json"
APP_NAME = "Custom Input"

enabled = False
current_preset = {}


def print_colored(text: str, color: str):
    stdout.write(colored(text, color) + "\n")
    stdout.flush()


def input_colored(text: str, color: str):
    stdout.write(colored(text, color))
    stdout.flush()
    return stdin.readline().replace("\n", "")


def messagebox(flag: int=MB_OK, text: str="Text"):
    MessageBox(GetConsoleWindow(), text, APP_NAME, flag)


def on_press(key_name: str):
    keyboard.press(current_preset[key_name])


def on_release(key_name: str):
    keyboard.release(current_preset[key_name])


def check_keys():
    keys = list(current_preset.keys()) + list(current_preset.values())

    for key in keys:
        try:
            keyboard.is_pressed(key)
        except ValueError:
            messagebox(MB_ICONERROR, f'"{key}" is not a valid key')
            exit()


def enable():
    global enabled

    for key, _ in current_preset.items():
        keyboard.on_press_key(key, lambda event: on_press(event.name.lower()), True)
        keyboard.on_release_key(key, lambda event: on_release(event.name.lower()), True)

    enabled = True


def disable():
    global enabled

    keyboard.unhook_all()
    enabled = False


def select_preset():
    global current_preset

    if not path.exists(PRESETS_FILE):
        messagebox(MB_ICONERROR, PRESETS_FILE + " not found")
        exit()

    print("AVAILABLE PRESETS:")

    with open(PRESETS_FILE) as f:
        presets = json.load(f)
        f.close()

    for i, preset_name in enumerate(presets):
        print_colored(f"{i + 1}. {preset_name}", "yellow")

    preset_name = input_colored("\npreset name: ", "cyan")
    system("cls")

    try:
        current_preset = presets[preset_name]
        check_keys()
        print_colored(f'running "{preset_name}" preset!', "green")
    except KeyError:
        messagebox(MB_ICONERROR, f'"{preset_name}" is not a valid preset')
        select_preset()


def main():
    SetConsoleTitle(APP_NAME)
    select_preset()
    check_keys()

    while True:
        caps_lock_state = GetKeyState(CAPS_LOCK_VK)

        if caps_lock_state == 1 and not enabled:
            enable()
        elif caps_lock_state == 0 and enabled:
            disable()

        time.sleep(REFRESH_RATE)


if __name__ == "__main__":
    main()
