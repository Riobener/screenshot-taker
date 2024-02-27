import tkinter as tk
import win32clipboard
import win32con
import win32gui
from PIL import ImageTk, ImageGrab
from ctypes import windll
from io import BytesIO

from utils.config_utils import get_save_path

GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080


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
        self.window.after(10, self.set_appwindow, self.window)
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
        windll.user32.SetWindowPos(self.window.winfo_id(), -1, self.window.winfo_x(), self.window.winfo_y(), 0, 0,
                                   0x0001)
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

    def set_appwindow(self, root):
        hwnd = windll.user32.GetParent(root.winfo_id())
        style = windll.user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
        style = style & ~WS_EX_TOOLWINDOW
        style = style | WS_EX_APPWINDOW
        windll.user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, style)
        root.withdraw()
        root.after(10, root.deiconify)
