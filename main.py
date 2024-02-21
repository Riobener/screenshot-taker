import os
import tkinter as tk
from io import BytesIO

import win32clipboard
import keyboard
import win32con
import win32gui
import sys
import yaml

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

    def take_screenshot(self, coords=None):
        self.screenshot = ImageGrab.grab()
        if coords is not None:
            self.screenshot = self.screenshot.crop(coords)
        return self

    def save(self):
        self.screenshot.save(get_save_path())
        self.send_to_clipboard()
        self.screenshot.close()

    def save_with_selecting_area(self):
        self.window = tk.Tk()
        self.window.attributes('-fullscreen', True)
        self.window.attributes('-topmost', True)
        self.window.wm_attributes("-topmost", 1)
        self.window.after(0, lambda: self.window.lift())
        self.window.resizable(False, False)
        self.window.focus_set()
        self.window.focus_force()
        self.window.overrideredirect(True)
        self.window.wm_attributes("-topmost", 1)
        win32gui.SetWindowPos(self.window.winfo_id(), win32con.HWND_TOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
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
                                                    outline='red')
        self.canvas.bind('<Button-1>', self.get_mouse_position)
        self.canvas.bind('<B1-Motion>', self.update_selected_area)
        self.canvas.bind('<ButtonRelease-1>', self.release_and_crop)
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
        return [left + 1, top + 1, right, bottom]

    def release_and_crop(self, event):
        self.crop_selected_image()
        self.window.destroy()

    def crop_selected_image(self):
        self.screenshot = self.screenshot.crop(self.get_converted_current_coords())
        self.save()

    def send_to_clipboard(self):
        output = BytesIO()
        self.screenshot.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()


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
    config = load_config()
    directory_path = config.get("directory_path")
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    filename = f"Screenshot_{current_time}.png"
    if directory_path == 'desktop':
        save_path = os.path.join(desktop_path, filename)
    else:
        save_path = os.path.join(directory_path, filename)
    return save_path


def on_selecting_area():
    ScreenTaker().take_screenshot().save_with_selecting_area()


def on_fullscreen():
    ScreenTaker().take_screenshot().save()


def on_foreground_window():
    window_coords = list(win32gui.GetWindowRect(GetForegroundWindow()))
    window_coords[0] = window_coords[0] + 8
    if window_coords[2] - 8 >= 0:
        window_coords[2] = window_coords[2] - 8
    window_coords[1] = window_coords[1] + 8
    if window_coords[3] - 10 >= 0:
        window_coords[3] = window_coords[3] - 10
    ScreenTaker().take_screenshot(window_coords).save()


def load_config():
    with open("config.yaml", 'r') as f:
        config = yaml.safe_load(f)
    return config


# Other method for keyboard
# HOTKEYS = [
#     {"keys": (win32con.VK_F12, win32con.MOD_CONTROL), "command": on_selecting_area},
#     {"keys": (win32con.VK_F11, win32con.MOD_CONTROL), "command": on_fullscreen},
#     {"keys": (win32con.VK_F10, win32con.MOD_CONTROL), "command": on_foreground_window},
# ]
# for i, h in enumerate(HOTKEYS):
#     vk, modifiers = h["keys"]
#     print("Registering id", i, "for key", vk)
#     if not ctypes.windll.user32.RegisterHotKey(None, i, modifiers, vk):
#         print("Unable to register id", i)
# try:
#     msg = wintypes.MSG()
#     while ctypes.windll.user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
#         if msg.message == win32con.WM_HOTKEY:
#             HOTKEYS[msg.wParam]["command"]()
#         ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
#         ctypes.windll.user32.DispatchMessageA(ctypes.byref(msg))
# finally:
#     for i, h in enumerate(HOTKEYS):
#         ctypes.windll.user32.UnregisterHotKey(None, i)

def main():
    config = load_config()
    fullscreen_hotkey = config.get("fullscreen_hotkey")
    selecting_area_hotkey = config.get("selecting_area_hotkey")
    foreground_screen_hotkey = config.get("foreground_screen_hotkey")
    keyboard.add_hotkey(selecting_area_hotkey, on_selecting_area)
    keyboard.add_hotkey(fullscreen_hotkey, on_fullscreen)
    keyboard.add_hotkey(foreground_screen_hotkey, on_foreground_window)
    PyStrayWorker()


if __name__ == "__main__":
    main()
