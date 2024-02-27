import ctypes
from ctypes import wintypes

import win32con


def register_and_listen_hotkeys(HOTKEYS):
    for i, h in enumerate(HOTKEYS):
        vk, modifiers = h["keys"]
        print("Registering id", i, "for key", vk)
        if not ctypes.windll.user32.RegisterHotKey(None, i, modifiers, vk):
            print("Unable to register id", i)
    try:
        msg = wintypes.MSG()
        while ctypes.windll.user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
            if msg.message == win32con.WM_HOTKEY:
                HOTKEYS[msg.wParam]["command"]()
            ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
            ctypes.windll.user32.DispatchMessageA(ctypes.byref(msg))
    finally:
        for i, h in enumerate(HOTKEYS):
            ctypes.windll.user32.UnregisterHotKey(None, i)
