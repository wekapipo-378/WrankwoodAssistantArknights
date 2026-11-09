"""状态机回归：按费用模型序列喂真实截图，验证全局帧。

序列（与计数器模型一致）：
  前 10 个周期：0,1,2..29（30 帧）
  第 11 个周期：0,1,2..29,30（31 帧）
  之后每个周期：1,2..29,30（30 帧）
"""

from __future__ import annotations

import sys

import cv2

sys.path.insert(0, ".")
from waa.core.frame_counter import WAAFrameCounter  # noqa: E402

CACHE: dict[str, "cv2.Mat"] = {}


def load(path: str):
    if path not in CACHE:
        CACHE[path] = cv2.cvtColor(cv2.imread(path), cv2.COLOR_BGR2GRAY)
    return CACHE[path]


def pre_img(phase: int):
    return load(f"out/bar_shots/pre/{phase + 1:02d}.png")


def post_img(phase: int):
    if phase == 30:
        return load("out/bar_shots/post/01.png")
    return load(f"out/bar_shots/post/{phase + 1:02d}.png")


def build_schedule():
    """返回 [(regime, phase, expected_frame)]，regime: pre/trans/post。"""
    out = []
    f = 0
    for _ in range(10):
        for p in [0, 1] + list(range(2, 30)):
            out.append(("pre", p, f))
            f += 1
    for p in [0, 1] + list(range(2, 30)) + [30]:
        out.append(("trans", p, f))
        f += 1
    for _ in range(2):
        for p in [1] + list(range(2, 30)) + [30]:
            out.append(("post", p, f))
            f += 1
    return out


def main() -> None:
    sched = build_schedule()
    counter = WAAFrameCounter()
    bad = 0
    for regime, ph, expected in sched:
        if ph == 30 or regime == "post":
            gray = post_img(ph)
        else:
            gray = pre_img(ph)
        counter.update(gray)
        if counter.frame != expected:
            print(f"MISMATCH {regime} phase={ph}: frame={counter.frame} "
                  f"expected={expected} warn={counter.warn}")
            bad += 1
            if bad > 5:
                break
    print(f"完整序列 {len(sched)} 帧: {'OK' if bad == 0 else 'FAIL'} "
          f"final_frame={counter.frame} expected={sched[-1][2]}")

    # 重复采样（暂停）不涨帧
    counter = WAAFrameCounter()
    for ph in [0, 1, 2, 3, 3, 3, 4]:
        counter.update(pre_img(ph))
    print(f"重复采样: frame={counter.frame} (期望 4) "
          f"{'OK' if counter.frame == 4 else 'FAIL'}")

    # 单周期内漏采一帧
    counter = WAAFrameCounter()
    for ph in [0, 1, 3]:
        counter.update(pre_img(ph))
    print(f"漏帧: frame={counter.frame} (期望 3) "
          f"{'OK' if counter.frame == 3 else 'FAIL'}")


if __name__ == "__main__":
    main()
