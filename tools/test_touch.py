"""合成触控测试：点暂停，或从干员卡拖到地格。

用法（管理员权限）：
    python tools\\test_touch.py --click-pause
    python tools\\test_touch.py --drag
"""

from __future__ import annotations

import argparse
import ctypes
import os
import sys
import time

import win32gui

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from waa.core.capture import grab_client  # noqa: E402
from waa.core.touch_input import TouchDevice  # noqa: E402
from PIL import Image  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "out")

PAUSE_RATIO = (0.945, 0.075)
CARD_RATIO = (0.74, 0.911)
TILE_RATIO = (0.4695, 0.4444)


def shot(hwnd: int, name: str) -> None:
    img = grab_client(hwnd)
    Image.fromarray(img[:, :, ::-1]).save(os.path.join(OUT, name))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--click-pause", action="store_true")
    parser.add_argument("--resume-first", action="store_true")
    parser.add_argument("--drag", action="store_true")
    parser.add_argument("--direction", action="store_true")
    parser.add_argument("--card", nargs=2, type=float, default=CARD_RATIO)
    parser.add_argument("--tile", nargs=2, type=float, default=TILE_RATIO)
    args = parser.parse_args()

    hwnd = win32gui.FindWindow("UnityWndClass", "明日方舟")
    if ctypes.windll.user32.IsZoomed(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(0.3)

    l, t, r, b = win32gui.GetClientRect(hwnd)
    w, h = r - l, b - t
    origin = win32gui.ClientToScreen(hwnd, (0, 0))

    def s(rx: float, ry: float) -> tuple[int, int]:
        return origin[0] + int(rx * w), origin[1] + int(ry * h)

    dev = TouchDevice()
    try:
        shot(hwnd, "touch_before.png")
        if args.click_pause:
            dev.click(*s(*PAUSE_RATIO))
            time.sleep(0.5)
            print("touch clicked pause")
        if args.resume_first:
            dev.click(*s(*PAUSE_RATIO))
            time.sleep(0.3)
            print("touch resumed")
        if args.drag:
            card = s(*args.card)
            middle = (card[0], int(card[1] - 0.03 * h))
            tile = s(*args.tile)
            points = [card]
            for i in range(1, 11):
                points.append((int(card[0] + (middle[0] - card[0]) * i / 10), int(card[1] + (middle[1] - card[1]) * i / 10)))
            for i in range(1, 21):
                points.append((int(middle[0] + (tile[0] - middle[0]) * i / 20), int(middle[1] + (tile[1] - middle[1]) * i / 20)))
            dev.drag(points)
            time.sleep(0.5)
            print("touch dragged")
        if args.direction:
            tile = s(*args.tile)
            target = (tile[0] + int(0.20 * w), tile[1])
            points = [tile]
            for i in range(1, 16):
                points.append((int(tile[0] + (target[0] - tile[0]) * i / 15), tile[1]))
            dev.drag(points)
            time.sleep(0.5)
            print("touch direction right")
        shot(hwnd, "touch_after.png")
    finally:
        dev.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
