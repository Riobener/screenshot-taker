import asyncio
import os
import tkinter as tk
import keyboard
import win32gui
import sys

from datetime import datetime
from PIL import ImageTk, ImageGrab, Image
from pystray import Icon as icon, Menu as menu, MenuItem as item
from win32gui import GetForegroundWindow
from win32api import GetSystemMetrics
WIDTH, HEIGHT = GetSystemMetrics(0), GetSystemMetrics(1)


class ScreenTaker:
    def __init__(self):
        self.rect_id = None
        self.canvas = None
        self.boty = None
        self.botx = None
        self.topy = None
        self.topx = None
        self.img = None
        self.window = None
        self.screenshot = None

    async def take_with_selecting_area(self):
        self.screenshot = ImageGrab.grab()
        self.window = tk.Tk()
        self.window.attributes('-fullscreen', True)
        self.window.attributes('-topmost', True)
        self.window.resizable(False, False)
        self.window.title("Select Area")
        self.window.configure(background='grey')
        self.img = ImageTk.PhotoImage(self.screenshot)
        self.canvas = tk.Canvas()
        self.topx, self.topy, self.botx, self.boty = 0, 0, 0, 0
        self.rect_id = None
        self.canvas = tk.Canvas(self.window, width=self.img.width(), height=self.img.height(),
                                borderwidth=0, highlightthickness=0)
        self.canvas.pack(expand=True)
        self.canvas.img = self.img
        self.canvas.create_image(0, 0, image=self.img, anchor=tk.NW)
        self.rect_id = self.canvas.create_rectangle(self.topx, self.topy, self.topx, self.topy, dash=(2, 2), fill='',
                                                    outline='white')
        self.canvas.bind('<Button-1>', self.get_mouse_position)
        self.canvas.bind('<B1-Motion>', self.update_selected_area)
        self.canvas.bind('<ButtonRelease-1>', self.release_and_save)
        self.screenshot.close()
        self.window.mainloop()

    def get_mouse_position(self, event):
        self.topx, self.topy = event.x, event.y

    def update_selected_area(self, event):
        self.botx, self.boty = event.x, event.y
        self.canvas.coords(self.rect_id, self.topx, self.topy, self.botx, self.boty)

    def get_converted_current_coords(self):
        left = min(self.topx, self.botx)
        top = min(self.topy, self.boty)
        right = max(self.topx, self.botx)
        bottom = max(self.topy, self.boty)
        return [left, top, right, bottom]

    def release_and_save(self, event):
        self.window.destroy()
        self.take_screenshot()

    def take_screenshot(self, coords=None):
        if self.img is None:
            if coords is None:
                screenshot = ImageGrab.grab()
            else:
                screenshot = ImageGrab.grab(coords)
        else:
            screenshot = ImageGrab.grab(self.get_converted_current_coords())
        screenshot.save(get_save_path())
        screenshot.close()


class PyStrayWorker:
    def __init__(self):
        self.icon = icon('test', Image.open("assets/icon.png"), menu=menu(
            item(text="Full Screen", action=on_fullscreen, default=True),
            item(text="Select Area", action=on_selecting_area),
            item('Exit', action=self.end_program)))
        self.icon.run()

    def end_program(self):
        self.icon.stop()
        sys.exit()


def get_save_path():
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    filename = f"screenshot_{current_time}.png"
    save_path = os.path.join(desktop_path, filename)
    return save_path


def on_selecting_area():
    asyncio.run(ScreenTaker().take_with_selecting_area())


def on_fullscreen():
    ScreenTaker().take_screenshot()


def on_foreground_window():
    window_coords = list(win32gui.GetWindowRect(GetForegroundWindow()))
    window_coords[0] = window_coords[0] + 8
    if window_coords[2] - 8 >= 0:
        window_coords[2] = window_coords[2] - 8
    window_coords[1] = window_coords[1] + 8
    if window_coords[3] - 10 >= 0:
        window_coords[3] = window_coords[3] - 10
    ScreenTaker().take_screenshot(window_coords)


def main():
    print("Hotkey 'Ctrl+F12' set to take a screenshot with area selecting tool")
    print("Hotkey 'Ctrl+F11' set to take a screenshot of full screen")
    print("Hotkey 'Ctrl+F10' set to take a screenshot of foreground window")
    keyboard.add_hotkey('ctrl+f12', on_selecting_area)
    keyboard.add_hotkey('ctrl+f11', on_fullscreen)
    keyboard.add_hotkey('ctrl+f10', on_foreground_window)
    PyStrayWorker()


if __name__ == "__main__":
    main()
