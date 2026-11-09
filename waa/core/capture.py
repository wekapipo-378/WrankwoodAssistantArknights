"""WGC 抓屏，并裁到游戏客户区。

WGC 抓的是整个窗口（含标题栏、边框），而 UI 坐标都是相对客户区的，所以这里
按 ClientToScreen 与 GetWindowRect 的差值裁掉非客户区。
"""

from __future__ import annotations

import ctypes
from ctypes import wintypes

import numpy as np
import win32gui

from windows_capture import WindowsCapture


def _enable_dpi_awareness() -> None:
    """让 GetClientRect 返回物理像素，和 WGC 抓到的画面尺寸一致。

    Windows 125% 缩放时，未声明 DPI 感知的进程会拿到 1536x864 的逻辑尺寸，
    而游戏实际画面是 1920x1080；不处理会把右边/下边裁掉（暂停键就在右边）。
    """
    try:
        # SetProcessDPIAware 返回 0 表示失败；必须检查返回值，
        # 否则会误以为已经设置成功（之前就是这个坑）。
        if ctypes.windll.user32.SetProcessDPIAware():
            return
    except Exception:
        pass
    try:
        if ctypes.windll.shcore.SetProcessDpiAwareness(1) == 0:
            return
    except Exception:
        pass


_enable_dpi_awareness()


def visible_window_rect(hwnd: int) -> tuple[int, int, int, int]:
    """DWM 可见窗口边界（不含 Win10/11 的透明外扩边框），单位是屏幕坐标。"""
    rect = wintypes.RECT()
    hr = ctypes.windll.dwmapi.DwmGetWindowAttribute(
        hwnd, 9, ctypes.byref(rect), ctypes.sizeof(rect)  # DWMWA_EXTENDED_FRAME_BOUNDS
    )
    if hr != 0:
        return win32gui.GetWindowRect(hwnd)
    return rect.left, rect.top, rect.right, rect.bottom


def client_rect(hwnd: int) -> tuple[int, int, int, int]:
    """返回客户区在屏幕上的 (left, top, width, height)。"""
    left, top = win32gui.ClientToScreen(hwnd, (0, 0))
    l, t, r, b = win32gui.GetClientRect(hwnd)
    return left, top, r - l, b - t


def client_offset(hwnd: int) -> tuple[int, int]:
    """客户区左上角相对于 WGC 画面（= DWM 可见边界）左上角的偏移。"""
    wl, wt, _, _ = visible_window_rect(hwnd)
    cl, ct = win32gui.ClientToScreen(hwnd, (0, 0))
    return cl - wl, ct - wt


def grab_client(hwnd: int) -> np.ndarray:
    """抓一帧客户区，返回 BGR numpy 数组。窗口被遮挡也能抓到。"""
    result: dict[str, np.ndarray] = {}

    capture = WindowsCapture(window_hwnd=hwnd, cursor_capture=False, draw_border=False)

    @capture.event
    def on_frame_arrived(frame, capture_control):
        try:
            # frame_buffer 是 BGRA；取前三个通道就是 BGR
            result["img"] = frame.frame_buffer[:, :, :3].copy()
        finally:
            capture_control.stop()

    @capture.event
    def on_closed():
        pass

    capture.start()
    if "img" not in result:
        raise RuntimeError("WGC 抓帧失败")

    img = result["img"]
    ox, oy = client_offset(hwnd)
    _, _, cw, ch = client_rect(hwnd)
    # 防御：窗口边框变化时不要把数组切越界
    x0 = max(0, min(ox, img.shape[1] - 1))
    y0 = max(0, min(oy, img.shape[0] - 1))
    x1 = min(img.shape[1], x0 + cw)
    y1 = min(img.shape[0], y0 + ch)
    return img[y0:y1, x0:x1].copy()
