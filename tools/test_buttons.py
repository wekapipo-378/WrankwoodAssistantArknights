"""验证 WAA 自己的按钮识别。"""

from __future__ import annotations

import glob
import os
import sys

import cv2

sys.path.insert(0, ".")
from waa.core.buttons import read_buttons  # noqa: E402

SESSION = [
    "_startup_check_039.png",
]

# 与 button_samples 保存顺序一致的原始文件映射
SAMPLES = sorted(glob.glob("out/button_samples/*.png"))
SAMPLES = [p for p in SAMPLES if "_montage" not in p]


def original_path(saved: str) -> str | None:
    name = os.path.basename(saved)
    stem = name
    # line05305_0_000.png -> out/session_images/line05305_0.jpg
    base = stem.rsplit("_", 1)[0]
    for ext in (".jpg", ".png"):
        p = os.path.join("out/session_images", base + ext)
        if os.path.exists(p):
            return p
    if base.startswith("_startup_check"):
        return os.path.join("out/bar_shots", "_startup_check.png")
    return None


def main() -> None:
    print("== 40 张历史截图 ==")
    for i, saved in enumerate(SAMPLES):
        p = original_path(saved)
        if not p:
            print(i, saved, "NO ORIGINAL")
            continue
        img = cv2.imread(p)
        if img is None:
            print(i, p, "READ FAIL")
            continue
        if img.shape[1] != 1920 or img.shape[0] != 1080:
            print(i, p, "SIZE", img.shape)
            continue
        st = read_buttons(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
        print(f"{i:02d} {os.path.basename(p):28s} "
              f"run={st.running} m={st.m} bullet={st.bullet} "
              f"pdrawn={st.pause_drawn} sdrawn={st.speed_drawn}")

    print("\n== 60 张 bar_shots（pre=1X 暂停，post=2X 暂停）==")
    bad = 0
    for phase, m_exp in (("pre", 1.0), ("post", 2.0)):
        for p in sorted(glob.glob(f"out/bar_shots/{phase}/*.png")):
            img = cv2.imread(p)
            st = read_buttons(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
            ok = (st.running is False and st.m == m_exp and not st.bullet)
            if not ok:
                bad += 1
                print("BAD", p, st)
    print(f"bar_shots: {60 - bad}/60 OK")


if __name__ == "__main__":
    main()
