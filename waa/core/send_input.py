"""SendInput 封装（MAA SeizeInput 的 Python 版）。

比 mouse_event 更现代，也是 MAA 在 PC 端实际使用的方式。
"""

from __future__ import annotations

import ctypes
import time
from ctypes import wintypes

user32 = ctypes.windll.user32

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1

MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_ABSOLUTE = 0x8000

KEYEVENTF_KEYUP = 0x0002

ULONG_PTR = ctypes.c_ulonglong


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD),
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUT_UNION)]


def _send(*inputs: INPUT) -> int:
    array = (INPUT * len(inputs))(*inputs)
    return user32.SendInput(len(inputs), array, ctypes.sizeof(INPUT))


def _mouse_input(flags: int, dx: int = 0, dy: int = 0, data: int = 0) -> INPUT:
    item = INPUT()
    item.type = INPUT_MOUSE
    item.u.mi = MOUSEINPUT(dx, dy, data, flags, 0, 0)
    return item


def set_cursor(x: int, y: int) -> None:
    user32.SetCursorPos(int(x), int(y))


def mouse_down(button: str = "left") -> None:
    flags = MOUSEEVENTF_LEFTDOWN if button == "left" else MOUSEEVENTF_RIGHTDOWN
    _send(_mouse_input(flags))
    time.sleep(0.02)


def mouse_up(button: str = "left") -> None:
    flags = MOUSEEVENTF_LEFTUP if button == "left" else MOUSEEVENTF_RIGHTUP
    _send(_mouse_input(flags))
    time.sleep(0.02)


def mouse_move_abs(screen_x: int, screen_y: int) -> None:
    sw = user32.GetSystemMetrics(0)
    sh = user32.GetSystemMetrics(1)
    nx = int(screen_x * 65535 / max(1, sw))
    ny = int(screen_y * 65535 / max(1, sh))
    _send(_mouse_input(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, nx, ny))


def click(x: int, y: int, button: str = "left") -> None:
    set_cursor(x, y)
    time.sleep(0.06)
    mouse_down(button)
    time.sleep(0.05)
    mouse_up(button)


def key_down(vk: int) -> None:
    item = INPUT()
    item.type = INPUT_KEYBOARD
    scan = user32.MapVirtualKeyW(vk, 0)
    item.u.ki = KEYBDINPUT(vk, scan, 0, 0, 0)
    _send(item)


def key_up(vk: int) -> None:
    item = INPUT()
    item.type = INPUT_KEYBOARD
    scan = user32.MapVirtualKeyW(vk, 0)
    item.u.ki = KEYBDINPUT(vk, scan, KEYEVENTF_KEYUP, 0, 0)
    _send(item)


def press_key(vk: int) -> None:
    key_down(vk)
    time.sleep(0.03)
    key_up(vk)
    time.sleep(0.05)
