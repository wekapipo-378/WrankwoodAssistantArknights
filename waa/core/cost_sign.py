"""局内负费：费用数字左边的减号。

坐标（1920x1080，UI 100%，真实截图实测）：
    减号框 x1790..1875, y768..779（在费用条上方那一行，覆盖数字整行）
    正常 100% 数字的最长白横杠实测 15px，减号 31px，阈值 21px 两边都有余量；
    框做宽是为了 -10 以上的多位数负费（减号会往左移）。

判据：
  - 框内 99 分位 < 220：这一小块根本没白，判“未知”（加载/菜单/被遮住）
  - 否则白门限 = max(220, 99分位 * 0.95)；减号芯会被 h264/缩放压到 240 左右，
    用 250 会整段漏判（实测踩过），所以门限跟着框内白电平走
  - 任意一行里连续白横杠 >= 21px 判负费；正常数字笔画实测最长不到 21px，
    减号在 1920x1080 实测 31px，余量足够

不读数字、不做 OCR，只看这条横杠。
"""

from __future__ import annotations

import numpy as np

SIGN_X0, SIGN_X1 = 1790, 1875
SIGN_Y0, SIGN_Y1 = 768, 779

SIGN_MIN_RUN = 21
SIGN_WHITE_FLOOR = 220.0
SIGN_WHITE_FRAC = 0.95


def _longest_run(row: np.ndarray) -> int:
    best = cur = 0
    for v in row:
        if v:
            cur += 1
            if cur > best:
                best = cur
        else:
            cur = 0
    return best


def sign_detail(gray: np.ndarray):
    """返回 (negative, longest_run, p99, threshold)。

    negative 为 None 表示这一小块读不出来（不是“没有减号”）。
    """
    roi = gray[SIGN_Y0:SIGN_Y1, SIGN_X0:SIGN_X1]
    if roi.size == 0:
        return None, 0, 0.0, 0.0
    p99 = float(np.percentile(roi, 99))
    if p99 < SIGN_WHITE_FLOOR:
        return None, 0, p99, 0.0
    threshold = max(SIGN_WHITE_FLOOR, p99 * SIGN_WHITE_FRAC)
    run = 0
    for row in roi > threshold:
        run = max(run, _longest_run(row))
    return run >= SIGN_MIN_RUN, run, p99, threshold


def read_sign(gray: np.ndarray) -> bool | None:
    """True=负费，False=正常，None=读不到（保持上一次状态）。"""
    return sign_detail(gray)[0]
