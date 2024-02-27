import asyncio
import os

import win32con
import win32gui
from PIL import Image
from pystray import Icon as icon, Menu as menu, MenuItem as item
from win32gui import GetForegroundWindow
from core.screen_taker import ScreenTaker
from utils.hotkey_utils import register_and_listen_hotkeys


class PyStrayWorker:
    def __init__(self):
        self.icon = icon('test', Image.open("assets/icon.png"), menu=menu(
            item(text="Full Screen", action=on_fullscreen, default=True),
            item(text="Select Area", action=on_selecting_area),
            item('Exit', action=self.end_program)))
        self.icon.run_detached()

    def end_program(self):
        self.icon.stop()
        os._exit(1)


def on_selecting_area():
    ScreenTaker().take_screenshot().save_with_selecting_area()


def on_fullscreen():
    ScreenTaker().take_screenshot().save()


def on_foreground_window():
    window_coords = list(win32gui.GetWindowRect(GetForegroundWindow()))
    window_coords[0] = window_coords[0] + 8
    if window_coords[2] - 8 >= 0:
        window_coords[2] = window_coords[2] - 8
    window_coords[1] = window_coords[1] + 4
    if window_coords[3] - 10 >= 0:
        window_coords[3] = window_coords[3] - 10
    ScreenTaker().take_screenshot(window_coords).save()


HOTKEYS = [
    {"keys": (win32con.VK_F12, win32con.MOD_CONTROL), "command": on_selecting_area},
    {"keys": (win32con.VK_F11, win32con.MOD_CONTROL), "command": on_fullscreen},
    {"keys": (win32con.VK_F10, win32con.MOD_CONTROL), "command": on_foreground_window},
]


async def main():
    asyncio.ensure_future(register_and_listen_hotkeys(HOTKEYS))
    PyStrayWorker()


if __name__ == "__main__":
    asyncio.run(main())
