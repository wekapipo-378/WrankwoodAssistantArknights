"""用 60 张真实截图验证分类器：pre/post 各应为 0/30..29/30。"""

from __future__ import annotations

import sys

import cv2

sys.path.insert(0, ".")
from waa.core.frame_counter import classify_frame  # noqa: E402

EXPECT_FILLS = [5, 12, 18, 24, 30, 37, 43, 49, 55, 62, 68, 74, 80, 86,
                93, 99, 105, 111, 118, 124, 130, 136, 143, 149, 155, 161,
                168, 174]


def main() -> None:
    bad = 0
    for phase_dir in ("pre", "post"):
        print(f"\n===== {phase_dir} =====")
        for i in range(1, 31):
            img = cv2.imread(f"out/bar_shots/{phase_dir}/{i:02d}.png")
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            v = classify_frame(gray)
            print(f"{i:02d}  w={v.width:3d}  cap={v.cap_diff:6.1f}  "
                  f"phase={v.phase:2d}  {v.kind}")

    # 期望序列
    pre_expected = [0, 1] + list(range(2, 30))
    post_expected = [30, 1] + list(range(2, 30))
    for phase_dir, expected in (("pre", pre_expected), ("post", post_expected)):
        got = []
        for i in range(1, 31):
            img = cv2.imread(f"out/bar_shots/{phase_dir}/{i:02d}.png")
            got.append(classify_frame(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)).phase)
        ok = got == expected
        print(f"\n{phase_dir}: {'OK' if ok else 'MISMATCH'}")
        print("  got     ", got)
        print("  expected", expected)
        if not ok:
            bad += 1
    print("\nRESULT:", "PASS" if bad == 0 else f"FAIL ({bad})")


if __name__ == "__main__":
    main()
