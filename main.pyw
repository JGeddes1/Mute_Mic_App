import os
import keyboard
import win32api
import win32gui
import pywintypes
from win10toast import ToastNotifier

WM_APPCOMMAND = 0x319
WM_BASE_HEX = 0xFFFF + 1
APPCOMMAND_MICROPHONE_VOLUME_DOWN = WM_BASE_HEX * 25
APPCOMMAND_MICROPHONE_VOLUME_UP = WM_BASE_HEX * 26

# Variable to track mute/unmute state
is_muted = False

toast = ToastNotifier()
toast.show_toast("Mute Unmute App", "The process has been started", duration = 3)
os.chdir("F:/Coding projects 2024/python/Mute Microphone App/venv/Mute_Mic_App")

def toggle_mic():
    global is_muted
    hwnd_active = win32gui.GetForegroundWindow()

    if is_muted:
        # Unmute the microphone
        for x in range(80):
            win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_UP)
        print("Microphone Unmuted")
    else:
        # Mute the microphone
        for x in range(80):
            win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_DOWN)
        print("Microphone Muted")

    # Toggle the mute state
    is_muted = not is_muted

print("Ctrl+M = toggle mute/unmute, End = quit")
keyboard.add_hotkey('ctrl+m', toggle_mic)
keyboard.wait('end')