"""WAA 自己的右上角按钮识别：暂停/运行、1X/2X、子弹时间（>>>）。

不用外部模板，全部按白色字形的几何结构判断：

  暂停键 ROI
    运行中：两根竖条 = 2 个连通块
    已暂停：一个右三角 = 1 个连通块

  倍速键 ROI
    1X  ："1"(窄) + "X"(宽 42) + 单个三角（宽约 29）
    2X  ："2"(宽 34) + "X" + 两个三角（合并后宽约 49）
    >>> ：没有文字；三个灰三角（可能合并成一个宽约 66 的块）
          子弹时间时整个 HUD 会变暗，所以用相对阈值，不依赖绝对亮度。

坐标按 1920×1080、UI 100% 实测。
"""

from __future__ import annotations

from dataclasses import dataclass

import cv2
import numpy as np

SPEED_ROI = (1600, 20, 1700, 125)   # x0, y0, x1, y1
PAUSE_ROI = (1770, 30, 1850, 95)

# AFA 自动开局暂停检测的「倍速按钮白色像素」区域（game_monitor.ahk
# SpeedButtonPositionColor）：0.845~0.8807 宽、0.0713~0.087 高，1920×1080 下即
# x 1622~1691、y 77~94。AFA 在这里找 0xffffff（容差 10）——倍速按钮的白色
# 文字/数字真正画出来的那一刻，比费用条 armed 早、又不像相对亮度那样在加载
# 画面误判。
SPEED_WHITE_ROI = (1622, 77, 1691, 94)
SPEED_WHITE_MIN = 245   # 0xffffff 容差 10 → 各通道 ≥ 245

REL_THR = 0.35     # 阈值 = median + 0.35 * (max - median)
MIN_AREA = 90      # 小于这个面积的连通块当噪声


@dataclass
class ButtonState:
    running: bool | None    # True=运行中，False=暂停，None=没画出来/不确定
    m: float                # 倍速 1.0 / 2.0（不确定时给 1.0）
    bullet: bool            # 是否子弹时间（>>>）
    pause_drawn: bool
    speed_drawn: bool


def _components(roi: np.ndarray) -> list[tuple[int, int, int, int, int]]:
    """返回 (x, y, w, h, area)；坐标相对 ROI。"""
    if roi.size == 0:
        return []
    med = float(np.median(roi))
    mx = float(roi.max())
    if mx - med < 20.0:
        return []
    thr = med + REL_THR * (mx - med)
    mask = (roi > thr).astype(np.uint8)
    n, _lab, stats, _cent = cv2.connectedComponentsWithStats(mask, 8)
    out = []
    for i in range(1, n):
        x, y, w, h, area = (int(v) for v in stats[i])
        if area >= MIN_AREA:
            out.append((x, y, w, h, area))
    return out


def _read_speed(gray: np.ndarray) -> tuple[float, bool, bool]:
    """返回 (m, bullet, drawn)。"""
    x0, y0, x1, y1 = SPEED_ROI
    roi = gray[y0:y1, x0:x1]
    comps = _components(roi)

    tall = [c for c in comps
            if c[3] >= 33 and c[2] >= 18
            and 0 <= c[0] <= 70 and 15 <= c[1] <= 45]        # 文字："1"/"2"/"X"
    if tall:
        digit_w = min(c[2] for c in tall)                    # 1 窄、2 宽
        arrow_w = max((c[2] for c in comps
                       if c[1] >= 60 and c[3] <= 32), default=0)
        m = 2.0 if (arrow_w >= 38 or digit_w >= 28) else 1.0
        return m, False, True

    # 没有文字：可能是子弹时间 >>>（三个三角，可能合并成一块）
    wide = [c for c in comps
            if 10 <= c[0] <= 45 and 35 <= c[1] <= 62
            and 45 <= c[2] <= 85 and 20 <= c[3] <= 32]
    mid = [c for c in comps
           if 10 <= c[0] <= 90 and 38 <= c[1] <= 80
           and 18 <= c[2] <= 40 and 18 <= c[3] <= 32]
    if wide or len(mid) >= 3:
        return 1.0, True, True
    return 1.0, False, False


def _read_pause(gray: np.ndarray) -> tuple[bool | None, bool]:
    """返回 (running, drawn)。"""
    x0, y0, x1, y1 = PAUSE_ROI
    roi = gray[y0:y1, x0:x1]
    comps = [c for c in _components(roi)
             if 30 <= c[3] <= 55 and 18 <= c[2] <= 75]
    bars = [c for c in comps if 20 <= c[2] <= 38 and 36 <= c[3] <= 55]
    tri = [c for c in comps if 42 <= c[2] <= 72 and 30 <= c[3] <= 50]
    if len(bars) >= 2:
        ys = [c[1] for c in bars]
        hs = [c[3] for c in bars]
        if max(ys) - min(ys) <= 8 and max(hs) - min(hs) <= 8:
            return True, True   # 两根竖条
    if len(tri) == 1 and len(comps) == 1:
        return False, True      # 单个三角
    return None, False


def read_buttons(gray: np.ndarray) -> ButtonState:
    m, bullet, speed_drawn = _read_speed(gray)
    running, pause_drawn = _read_pause(gray)
    return ButtonState(running=running, m=m, bullet=bullet,
                       pause_drawn=pause_drawn, speed_drawn=speed_drawn)


def speed_white_pixel(gray: np.ndarray) -> bool:
    """AFA 开局暂停的精确信号：倍速按钮区域里有没有「纯白像素」。

    复刻 AFA game_monitor 的 `SafePixelSearch(倍速按钮区域, 0xffffff, 容差10)`：
    在 SPEED_WHITE_ROI 内找亮度 ≥ SPEED_WHITE_MIN 的像素。倍速按钮的白色文字/
    数字真正画出来时才会命中；加载/intro 阶段（画面整体偏暗、只有相对亮块）
    不会命中——这正是 WAA 的 speed_drawn（相对亮度阈值）误判的地方。
    """
    x0, y0, x1, y1 = SPEED_WHITE_ROI
    if gray.shape[0] < y1 or gray.shape[1] < x1:
        return False
    roi = gray[y0:y1, x0:x1]
    if roi.size == 0:
        return False
    return bool(np.any(roi >= SPEED_WHITE_MIN))
