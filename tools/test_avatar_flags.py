"""离线测试：MAA 旗标找卡 + 卡片头像框，适配头像缩小。"""

from __future__ import annotations

import os
import sys

import cv2
import numpy as np

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa.core.vision import (  # noqa: E402
    FLAG_TEMPLATE_PATH, OPER_AVATAR_MOVE, OPER_CLICK_MOVE,
    STANDARD_FLAG_PITCH, find_operator_flags, get_avatars, locate_avatar,
    operator_cards, rightmost_card_center)


def check(cond: bool, label: str) -> bool:
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def make_frame(flag_xs, target_x, avatar):
    rng = np.random.default_rng(20261104)
    frame = rng.integers(0, 255, (720, 1280, 3), dtype=np.uint8)
    flag = cv2.imread(FLAG_TEMPLATE_PATH, cv2.IMREAD_GRAYSCALE)
    flag_bgr = cv2.cvtColor(flag, cv2.COLOR_GRAY2BGR)
    pitch = (float(np.median(np.diff(flag_xs)))
             if len(flag_xs) >= 2 else STANDARD_FLAG_PITCH)
    scale = max(0.50, min(1.10, pitch / STANDARD_FLAG_PITCH))
    for fx in flag_xs:
        frame[592:592 + flag.shape[0], fx:fx + flag.shape[1]] = flag_bgr
    cx0 = target_x + OPER_CLICK_MOVE[0] * scale
    cy0 = 592 + OPER_CLICK_MOVE[1] * scale
    ax = int(round(cx0 + OPER_AVATAR_MOVE[0] * scale))
    ay = int(round(cy0 + OPER_AVATAR_MOVE[1] * scale))
    aw = int(round(OPER_AVATAR_MOVE[2] * scale))
    ah = int(round(OPER_AVATAR_MOVE[3] * scale))
    tpl = cv2.resize(avatar, (aw, ah), interpolation=cv2.INTER_AREA)
    frame[ay:ay + ah, ax:ax + aw] = cv2.cvtColor(tpl, cv2.COLOR_GRAY2BGR)
    want_x = (cx0 + OPER_CLICK_MOVE[2] * scale / 2) / 1280
    want_y = (cy0 + OPER_CLICK_MOVE[3] * scale / 2) / 720
    return frame, scale, want_x, want_y


def main() -> int:
    checks = []
    avatar = get_avatars("斑点")[0]

    # 1) 一张卡：旗标 -> 卡片 -> 头像框
    frame, scale, want_x, want_y = make_frame([800], 800, avatar)
    flags = find_operator_flags(frame)
    cards = operator_cards(frame)
    checks.append(check(len(flags) == 1 and len(cards) == 1,
                        f"单卡旗标/卡片 flags={len(flags)} cards={len(cards)}"))
    x, y, score = locate_avatar(frame, "斑点")
    checks.append(check(abs(x - want_x) < 0.004 and abs(y - want_y) < 0.004
                        and score >= 0.99,
                        f"单卡定位 ({x:.3f},{y:.3f}) want "
                        f"({want_x:.3f},{want_y:.3f}) score={score:.3f}"))

    # 2) 十张卡、间距 80px（头像缩小）：用最右卡作为目标
    flag_xs = [260 + 80 * i for i in range(10)]
    target = flag_xs[-1]
    frame, scale, want_x, want_y = make_frame(flag_xs, target, avatar)
    flags = find_operator_flags(frame)
    cards = operator_cards(frame)
    checks.append(check(len(flags) == 10 and len(cards) == 10,
                        f"十卡旗标 flags={len(flags)} cards={len(cards)}"))
    checks.append(check(abs(scale - 80 / STANDARD_FLAG_PITCH) < 0.01,
                        f"缩小比例 scale={scale:.3f}"))
    x, y, score = locate_avatar(frame, "斑点")
    checks.append(check(abs(x - want_x) < 0.004 and abs(y - want_y) < 0.004
                        and score >= 0.99,
                        f"缩小卡定位 ({x:.3f},{y:.3f}) want "
                        f"({want_x:.3f},{want_y:.3f}) score={score:.3f}"))
    pt = rightmost_card_center(frame)
    checks.append(check(pt is not None and abs(pt[0] - want_x) < 0.01,
                        f"最右卡 {pt} want_x={want_x:.3f}"))

    # 3) 真实截图（如果存在）：5 张卡，旗标应找到 5 个
    shot = os.path.join(ROOT, "out", "deploy_0_before.png")
    if os.path.exists(shot):
        img = cv2.imread(shot)
        flags = find_operator_flags(img)
        if not flags:
            print("SKIP 真实截图夹具不是战斗内画面（没有待部署旗标）")
        else:
            checks.append(check(len(flags) == 5,
                                f"真实截图旗标数 = {len(flags)}"))
            pt = rightmost_card_center(img)
            checks.append(check(pt is not None and 0.90 < pt[0] < 1.0,
                                f"真实截图最右卡 = {pt}"))

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
