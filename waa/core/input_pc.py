"""明日方舟 PC 端的输入层。

MAA 的注释里写得很清楚：PC 端按「真实光标位置」取坐标，不认消息里带的坐标。
所以这里采用和 MAA `PostMessageWithCursorPos` 一样的做法：
    1. 保存当前光标位置
    2. SetCursorPos 到目标位置（真实光标）
    3. PostMessage 鼠标消息
    4. 还原光标

打轴期间不要碰鼠标，否则会干扰点击位置。
"""

from __future__ import annotations

import ctypes
import time

import win32api
import win32con
import win32gui

# 让进程感知 DPI，避免高缩放显示器上客户区坐标和屏幕坐标对不上
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

_user32 = ctypes.windll.user32


def set_cursor_pos(x: int, y: int) -> None:
    """用 ctypes 直接调 SetCursorPos。

    pywin32 的 win32api.SetCursorPos 会拿残留的 GetLastError 当失败抛异常，
    实测光标其实已经移动了。这里忽略返回值，只看实际效果。
    """
    _user32.SetCursorPos(int(x), int(y))


def _make_lparam(x: int, y: int) -> int:
    return (y << 16) | (x & 0xFFFF)


class PCInput:
    def __init__(self, hwnd: int, hold_ms: int = 40, move_settle_ms: int = 8):
        self.hwnd = hwnd
        self.hold_ms = hold_ms
        self.move_settle_ms = move_settle_ms
        self._restore_cursor = True

    # ---- 基础 ----
    def client_size(self) -> tuple[int, int]:
        left, top, right, bottom = win32gui.GetClientRect(self.hwnd)
        return right - left, bottom - top

    def _client_to_screen(self, x: int, y: int) -> tuple[int, int]:
        return win32gui.ClientToScreen(self.hwnd, (x, y))

    def _post(self, msg: int, wparam: int, lparam: int) -> None:
        win32gui.PostMessage(self.hwnd, msg, wparam, lparam)

    def _move_cursor_to_client(self, x: int, y: int) -> None:
        sx, sy = self._client_to_screen(x, y)
        set_cursor_pos(sx, sy)
        if self.move_settle_ms:
            time.sleep(self.move_settle_ms / 1000.0)

    def _with_cursor_saved(self, fn):
        old = win32api.GetCursorPos()
        try:
            return fn()
        finally:
            if self._restore_cursor:
                set_cursor_pos(*old)

    # ---- 鼠标 ----
    def click_client(self, x: int, y: int) -> None:
        def action() -> None:
            self._move_cursor_to_client(x, y)
            self._post(win32con.WM_MOUSEMOVE, 0, _make_lparam(x, y))
            time.sleep(0.005)
            self._post(win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, _make_lparam(x, y))
            time.sleep(self.hold_ms / 1000.0)
            self._post(win32con.WM_LBUTTONUP, 0, _make_lparam(x, y))
            time.sleep(0.01)

        self._with_cursor_saved(action)

    def click_ratio(self, rx: float, ry: float) -> None:
        w, h = self.client_size()
        self.click_client(int(rx * w), int(ry * h))

    def move_client(self, x: int, y: int) -> None:
        def action() -> None:
            self._move_cursor_to_client(x, y)
            self._post(win32con.WM_MOUSEMOVE, 0, _make_lparam(x, y))
            time.sleep(0.05)

        self._with_cursor_saved(action)

    # ---- 前台真实输入（需要管理员权限）----
    def foreground(self) -> None:
        # 只恢复，不要最大化。SW_MAXIMIZE(3) 会把 1280x720 变成 1280x800，
        # 16:9 的 UI 坐标全部失效。
        if ctypes.windll.user32.IsZoomed(self.hwnd):
            win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
        try:
            win32gui.SetForegroundWindow(self.hwnd)
        except Exception:
            pass
        time.sleep(0.25)

    def click_client_real(self, x: int, y: int, hold_ms: int = 50) -> None:
        self._move_cursor_to_client(x, y)
        time.sleep(0.06)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(hold_ms / 1000.0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.05)

    def _move_cursor_event_abs(self, x: int, y: int) -> None:
        """用 mouse_event 绝对移动，确保游戏收到鼠标移动事件。

        只用 SetCursorPos 时，光标位置变了，但部分 Unity 游戏收不到 WM_MOUSEMOVE，
        拖拽会被当成一次点击。
        """
        sx, sy = self._client_to_screen(x, y)
        sw = win32api.GetSystemMetrics(0)
        sh = win32api.GetSystemMetrics(1)
        nx = int(sx * 65535 / max(1, sw))
        ny = int(sy * 65535 / max(1, sh))
        win32api.mouse_event(
            win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny, 0, 0
        )

    def drag_path_real(
        self,
        points: list[tuple[int, int]],
        duration_ms: int = 400,
        steps_per_segment: int = 12,
    ) -> None:
        if len(points) < 2:
            raise ValueError("drag_path_real 至少需要两个点")
        self._move_cursor_event_abs(*points[0])
        time.sleep(0.06)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(self.hold_ms / 1000.0)

        segment_ms = duration_ms / max(1, len(points) - 1)
        for (x1, y1), (x2, y2) in zip(points, points[1:]):
            interval = max(0.004, segment_ms / 1000.0 / steps_per_segment)
            for i in range(1, steps_per_segment + 1):
                t = i / steps_per_segment
                self._move_cursor_event_abs(int(x1 + (x2 - x1) * t), int(y1 + (y2 - y1) * t))
                time.sleep(interval)

        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.05)

    def drag_ratio_real(
        self,
        points: list[tuple[float, float]],
        duration_ms: int = 400,
    ) -> None:
        w, h = self.client_size()
        self.drag_path_real([(int(x * w), int(y * h)) for x, y in points], duration_ms)

    def press_esc_real(self) -> None:
        scan = ctypes.windll.user32.MapVirtualKeyW(win32con.VK_ESCAPE, 0)
        win32api.keybd_event(win32con.VK_ESCAPE, scan, 0, 0)
        time.sleep(0.04)
        win32api.keybd_event(win32con.VK_ESCAPE, scan, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.08)

    def swipe_client(
        self,
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        duration_ms: int = 300,
        steps: int = 20,
    ) -> None:
        def action() -> None:
            self._move_cursor_to_client(x1, y1)
            self._post(win32con.WM_MOUSEMOVE, 0, _make_lparam(x1, y1))
            time.sleep(0.01)
            self._post(win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, _make_lparam(x1, y1))
            time.sleep(self.hold_ms / 1000.0)

            interval = max(0.001, duration_ms / 1000.0 / max(1, steps))
            for i in range(1, steps + 1):
                t = i / steps
                x = int(x1 + (x2 - x1) * t)
                y = int(y1 + (y2 - y1) * t)
                self._move_cursor_to_client(x, y)
                self._post(win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON, _make_lparam(x, y))
                time.sleep(interval)

            self._post(win32con.WM_LBUTTONUP, 0, _make_lparam(x2, y2))
            time.sleep(0.01)

        self._with_cursor_saved(action)

    def swipe_ratio(
        self,
        r1: tuple[float, float],
        r2: tuple[float, float],
        duration_ms: int = 300,
    ) -> None:
        w, h = self.client_size()
        self.swipe_client(
            int(r1[0] * w), int(r1[1] * h),
            int(r2[0] * w), int(r2[1] * h),
            duration_ms,
        )

    # ---- 键盘 ----
    def press_esc(self) -> None:
        self._post(win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
        time.sleep(0.02)
        self._post(win32con.WM_KEYUP, win32con.VK_ESCAPE, 0xC0000001)
        time.sleep(0.02)
