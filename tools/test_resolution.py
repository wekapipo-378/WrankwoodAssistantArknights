"""离线测试：16:9 其他分辨率归一化到 1920×1080。"""

from __future__ import annotations

import os
import sys

import cv2
import numpy as np

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa.core.frame_counter import (  # noqa: E402
    classify_frame, normalize_frame,
)


def check(cond, label):
    print(("PASS " if cond else "FAIL ") + label)
    return bool(cond)


def main() -> int:
    checks = []
    for shape in ((720, 1280), (900, 1600), (1440, 2560)):
        img = np.zeros(shape, dtype=np.uint8)
        out = normalize_frame(img)
        checks.append(check(out.shape == (1080, 1920),
                            f"{shape} -> {out.shape}"))

    path = os.path.join(ROOT, "out", "bar_shots", "pre", "01.png")
    if os.path.exists(path):
        gray = cv2.cvtColor(cv2.imread(path), cv2.COLOR_BGR2GRAY)
        base = classify_frame(gray)
        for size in ((720, 1280), (900, 1600), (1440, 2560)):
            small = cv2.resize(gray, (size[1], size[0]),
                               interpolation=cv2.INTER_AREA)
            back = normalize_frame(small)
            got = classify_frame(back)
            same = (abs(got.width - base.width) <= 2
                    and got.kind == base.kind)
            checks.append(check(same,
                                f"{size} 费用条 {base.width}/{base.kind}"
                                f" -> {got.width}/{got.kind}"))
    else:
        print("SKIP 没有 out/bar_shots/pre/01.png")

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
