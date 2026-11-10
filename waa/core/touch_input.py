"""合成触控输入（MAA AnchoredTouch 的精简版）。

明日方舟 PC 端的拖拽只认触控事件，普通鼠标拖拽会被忽略。
这里用 Windows 的 CreateSyntheticPointerDevice / InjectSyntheticPointerInput
直接注入触控点，坐标是屏幕坐标。
"""

from __future__ import annotations

import ctypes
import time
from ctypes import wintypes

user32 = ctypes.windll.user32

PT_TOUCH = 2
POINTER_FEEDBACK_NONE = 3

POINTER_FLAG_INRANGE = 0x00000002
POINTER_FLAG_INCONTACT = 0x00000004
POINTER_FLAG_CONFIDENCE = 0x00000400
POINTER_FLAG_DOWN = 0x00010000
POINTER_FLAG_UPDATE = 0x00020000
POINTER_FLAG_UP = 0x00040000

DOWN_FLAGS = POINTER_FLAG_DOWN | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT | POINTER_FLAG_CONFIDENCE
UPDATE_FLAGS = POINTER_FLAG_UPDATE | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT | POINTER_FLAG_CONFIDENCE
UP_FLAGS = POINTER_FLAG_UP

TOUCH_MASK = 0x1 | 0x2 | 0x4  # CONTACTAREA | ORIENTATION | PRESSURE


class POINTER_INFO(ctypes.Structure):
    _fields_ = [
        ("pointerType", wintypes.DWORD),
        ("pointerId", wintypes.UINT),
        ("frameId", wintypes.UINT),
        ("pointerFlags", wintypes.DWORD),
        ("sourceDevice", wintypes.HANDLE),
        ("hwndTarget", wintypes.HWND),
        ("ptPixelLocation", wintypes.POINT),
        ("ptHimetricLocation", wintypes.POINT),
        ("ptPixelLocationRaw", wintypes.POINT),
        ("ptHimetricLocationRaw", wintypes.POINT),
        ("dwTime", wintypes.DWORD),
        ("historyCount", wintypes.UINT),
        ("InputData", wintypes.INT),
        ("dwKeyStates", wintypes.DWORD),
        ("PerformanceCount", ctypes.c_uint64),
        ("ButtonChangeType", wintypes.INT),
    ]


class POINTER_TOUCH_INFO(ctypes.Structure):
    _fields_ = [
        ("pointerInfo", POINTER_INFO),
        ("touchFlags", wintypes.DWORD),
        ("touchMask", wintypes.DWORD),
        ("rcContact", wintypes.RECT),
        ("rcContactRaw", wintypes.RECT),
        ("orientation", wintypes.UINT),
        ("pressure", wintypes.UINT),
    ]


class POINTER_PEN_INFO(ctypes.Structure):
    # 只是给 union 占位；实际只用 touchInfo
    _fields_ = [("raw", ctypes.c_byte * 200)]


class POINTER_TYPE_INFO(ctypes.Structure):
    class _U(ctypes.Union):
        _fields_ = [("touchInfo", POINTER_TOUCH_INFO), ("penInfo", POINTER_PEN_INFO)]

    _fields_ = [("type", wintypes.DWORD), ("u", _U)]


user32.CreateSyntheticPointerDevice.argtypes = [wintypes.DWORD, wintypes.ULONG, wintypes.DWORD]
user32.CreateSyntheticPointerDevice.restype = wintypes.HANDLE
user32.InjectSyntheticPointerInput.argtypes = [
    wintypes.HANDLE,
    ctypes.POINTER(POINTER_TYPE_INFO),
    wintypes.UINT,
]
user32.InjectSyntheticPointerInput.restype = wintypes.BOOL
user32.DestroySyntheticPointerDevice.argtypes = [wintypes.HANDLE]
user32.DestroySyntheticPointerDevice.restype = None


class TouchDevice:
    def __init__(self, pointer_id: int = 1, max_contacts: int = 11):
        self.pointer_id = pointer_id
        self.device = user32.CreateSyntheticPointerDevice(PT_TOUCH, max_contacts, POINTER_FEEDBACK_NONE)
        if not self.device:
            raise RuntimeError(f"CreateSyntheticPointerDevice failed: {ctypes.GetLastError()}")

    def close(self) -> None:
        if self.device:
            user32.DestroySyntheticPointerDevice(self.device)
            self.device = None

    def inject(self, screen_x: int, screen_y: int, flags: int) -> bool:
        info = POINTER_TYPE_INFO()
        info.type = PT_TOUCH
        ti = info.u.touchInfo
        ti.pointerInfo.pointerType = PT_TOUCH
        ti.pointerInfo.pointerId = self.pointer_id
        ti.pointerInfo.pointerFlags = flags
        ti.pointerInfo.ptPixelLocation.x = int(screen_x)
        ti.pointerInfo.ptPixelLocation.y = int(screen_y)
        ti.touchFlags = 0
        ti.touchMask = TOUCH_MASK
        ti.rcContact.left = int(screen_x) - 2
        ti.rcContact.top = int(screen_y) - 2
        ti.rcContact.right = int(screen_x) + 2
        ti.rcContact.bottom = int(screen_y) + 2
        ti.rcContactRaw = ti.rcContact
        ti.orientation = 90
        ti.pressure = 32000

        arr = (POINTER_TYPE_INFO * 1)(info)
        ok = user32.InjectSyntheticPointerInput(self.device, arr, 1)
        if not ok:
            raise RuntimeError(f"InjectSyntheticPointerInput failed: {ctypes.GetLastError()}")
        return True

    def touch_down(self, screen_x: int, screen_y: int) -> None:
        self.inject(screen_x, screen_y, DOWN_FLAGS)

    def touch_move(self, screen_x: int, screen_y: int) -> None:
        self.inject(screen_x, screen_y, UPDATE_FLAGS)

    def touch_up(self, screen_x: int, screen_y: int) -> None:
        self.inject(screen_x, screen_y, UP_FLAGS)

    def click(self, screen_x: int, screen_y: int, hold_ms: int = 50) -> None:
        self.touch_down(screen_x, screen_y)
        time.sleep(hold_ms / 1000.0)
        self.touch_up(screen_x, screen_y)
        time.sleep(0.05)

    def tap(self, screen_x: int, screen_y: int) -> None:
        """快速点一下，不在按下和抬起之间等待。

        AFA 的“暂停选中”要在一个逻辑帧内连点三次，普通 click 的 50ms
        按住会拖到下一帧，所以这里单独做一个几乎瞬时的 tap。
        """
        self.touch_down(screen_x, screen_y)
        self.inject(screen_x, screen_y, UPDATE_FLAGS)
        self.touch_up(screen_x, screen_y)

    def drag(self, points: list[tuple[int, int]], step_ms: int = 12) -> None:
        if len(points) < 2:
            raise ValueError("drag 至少需要两个点")
        x, y = points[0]
        self.touch_down(x, y)
        time.sleep(step_ms / 1000.0)
        for x, y in points[1:]:
            self.touch_move(x, y)
            time.sleep(step_ms / 1000.0)
        self.touch_up(x, y)
        time.sleep(0.05)
