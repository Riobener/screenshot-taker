import os
import tkinter as tk
import keyboard
import win32gui

from datetime import datetime
from PIL import ImageTk, ImageGrab, Image
from pystray import Icon as icon, Menu as menu, MenuItem as item
from win32gui import GetForegroundWindow

WIDTH, HEIGHT = 1920, 1080


class ScreenTaker:
    def __init__(self):
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

    def take_with_area(self):
        self.canvas = tk.Canvas(self.window, width=self.img.width(), height=self.img.height(),
                                borderwidth=0, highlightthickness=0)
        self.canvas.pack(expand=True)
        self.canvas.img = self.img
        self.canvas.create_image(0, 0, image=self.img, anchor=tk.NW)
        self.rect_id = self.canvas.create_rectangle(self.topx, self.topy, self.topx, self.topy, dash=(2, 2), fill='',
                                                    outline='white')
        self.canvas.bind('<Button-1>', self.get_mouse_position)
        self.canvas.bind('<B1-Motion>', self.update_selected_rectangle)
        self.canvas.bind('<ButtonRelease-1>', self.release_and_save)
        self.window.mainloop()

    def get_mouse_position(self, event):
        self.topx, self.topy = event.x, event.y

    def update_selected_rectangle(self, event):
        self.botx, self.boty = event.x, event.y
        self.canvas.coords(self.rect_id, self.topx, self.topy, self.botx, self.boty)

    def convert_coords(self):
        left = min(self.topx, self.botx)
        top = min(self.topy, self.boty)
        right = max(self.topx, self.botx)
        bottom = max(self.topy, self.boty)
        return [left, top, right, bottom]

    def release_and_save(self, event):
        self.window.destroy()
        self.take_by_coords(self.convert_coords())

    def take_by_coords(self, coords):
        self.screenshot = ImageGrab.grab(bbox=coords)
        self.screenshot.save(get_save_path())
        self.screenshot.close()

    def take_fullscreen(self):
        self.screenshot.save(get_save_path())
        self.screenshot.close()


class PyStrayWorker:
    def __init__(self):
        self.icon = icon('test', Image.open("icon.png"), menu=menu(
            item(text="Left-Click-Action", action=on_selecting_area, default=True),
            item(
                'Exit',
                self.end_program))).run()

    def end_program(self):
        self.icon.stop()
        exit(1)


def get_save_path():
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    filename = f"screenshot_{current_time}.png"
    save_path = os.path.join(desktop_path, filename)
    return save_path


def on_selecting_area():
    ScreenTaker().take_with_area()


def on_fullscreen():
    ScreenTaker().take_fullscreen()


def on_foreground_window():
    window_coords = list(win32gui.GetWindowRect(GetForegroundWindow()))
    print(window_coords)
    window_coords[0] = window_coords[0] + 8
    window_coords[2] = window_coords[2] - 8
    window_coords[1] = window_coords[1] + 8
    window_coords[3] = window_coords[3] - 10
    ScreenTaker().take_by_coords(window_coords)


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
