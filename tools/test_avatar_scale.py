"""离线测试：待部署区头像缩小时还能不能找到干员。"""

from __future__ import annotations

import os
import sys

import cv2
import numpy as np

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa.core.vision import get_avatars, locate_avatar  # noqa: E402


def check(cond: bool, label: str) -> bool:
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def main() -> int:
    base = get_avatars("斑点")[0]
    checks = []
    rng = np.random.default_rng(20261104)
    for scale in (1.0, 0.88, 0.80, 0.72, 0.60):
        tpl = (base if scale == 1.0 else cv2.resize(
            base, None, fx=scale, fy=scale,
            interpolation=cv2.INTER_AREA))
        frame = rng.integers(0, 255, (720, 1280, 3), dtype=np.uint8)
        x0, y0 = 760, 620
        frame[y0:y0 + tpl.shape[0], x0:x0 + tpl.shape[1]] = \
            cv2.cvtColor(tpl, cv2.COLOR_GRAY2BGR)
        x, y, score = locate_avatar(frame, "斑点")
        want_x = (x0 + tpl.shape[1] / 2) / 1280
        want_y = (y0 + tpl.shape[0] / 2) / 720
        checks.append(check(
            abs(x - want_x) < 0.004 and abs(y - want_y) < 0.004
            and score >= 0.99,
            f"scale={scale:.2f} -> ({x:.3f},{y:.3f}) "
            f"want ({want_x:.3f},{want_y:.3f}) score={score:.3f}"))
    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
