"""合成测试：局内负费在周期中途打开/关闭，帧号必须连续。

按 fp24 模型自己画费用条 + 数字笔画 + 减号，喂给 WAAFrameCounter。
"""

from __future__ import annotations

import os
import sys

import numpy as np

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa.core.cost_model import (  # noqa: E402
    EPS, SCALE, SPEED, measured_width, raw_width,
)
from waa.core.frame_counter import WAAFrameCounter  # noqa: E402


def render(width: int, cap_bright: bool, neg: bool) -> np.ndarray:
    """1920x1080 灰度图：只画费用条 ROI + 数字笔画 + 可选减号。"""
    g = np.full((1080, 1920), 50, np.uint8)
    # 参考底板 / 条帽
    g[811:817, 1728:1733] = 100
    g[811:817, 1737:1739] = 120 if cap_bright else 95
    # 条体
    if width >= 178:
        g[811:817, 1740:1920] = 255
    elif width > 0:
        g[811:817, 1740:1740 + width] = 255
    # 数字笔画（正常画面里框内的白横杠只有 6px，不能误报）
    g[760:790, 1845:1851] = 255
    if neg:
        # 减号：31px 纯白横杠
        g[772:776, 1824:1855] = 255
    return g


def build(rate=100.0, n=420, switch_on=95, switch_off=260):
    required = SCALE
    eps = EPS
    normal = max(1, int(round(SPEED * rate / 100.0)))
    neg_speed = max(1, normal // 2)
    acc = 0
    frames = []
    states = []
    for i in range(n):
        neg = switch_on <= i < switch_off
        speed = neg_speed if neg else normal
        w = measured_width(raw_width(acc / float(required)))
        nxt = acc + speed
        if nxt + eps >= required:
            nxt -= required
        nw = measured_width(raw_width(nxt / float(required)))
        cap = (w == 0 and nw > 0)
        frames.append(render(w, cap, neg))
        states.append((w, cap, neg))
        acc += speed
        if acc + eps >= required:
            acc -= required
    return frames, states


def main():
    counter = WAAFrameCounter(100.0)
    frames, states = build()
    deltas = []
    bad = 0
    for i, img in enumerate(frames):
        v = counter.update(img)
        deltas.append(counter.last_delta)
        if i == 0:
            if counter.frame != 0:
                print("首帧不是 0:", counter.frame)
                bad += 1
            continue
        # 允许 0（重复画面）或 1/2（从 0 宽暗段补帧），但累计帧必须对
        if counter.last_delta != 1:
            print(f"帧 {i}: true={i} counter={counter.frame} "
                  f"delta={counter.last_delta} state={states[i]} "
                  f"warn={counter.warn}")
        if counter.last_delta not in (0, 1, 2):
            print(f"帧 {i}: delta={counter.last_delta} warn={counter.warn}")
            bad += 1
    print(f"合成 {len(frames)} 帧: final={counter.frame} expected={len(frames)-1} "
          f"neg={counter.neg} speed={counter.speed}")
    print("deltas 统计: 0=", deltas.count(0), " 1=", deltas.count(1),
          " 2=", deltas.count(2), " warn=", sum(1 for _ in []))
    ok = bad == 0 and counter.frame == len(frames) - 1
    print("结果:", "OK" if ok else "FAIL")

    # 50% 战前 + 局内负费 = 120 帧/费（条每帧约 1.56px）
    c2 = WAAFrameCounter(50.0, neg_override=True)
    frames2, _ = build(rate=50.0, n=400, switch_on=0, switch_off=10 ** 9)
    lags = []
    for i, img in enumerate(frames2):
        c2.update(img)
        if i > 0 and c2.frame != i:
            lags.append((i, c2.frame, c2.last_delta, c2.uncertainty))
    print(f"50%+负费 120帧/费: final={c2.frame} expected={len(frames2)-1} "
          f"滞后拍数={len(lags)}")
    print("  前 12 个滞后:", lags[:12])
    ok2 = c2.frame == len(frames2) - 1
    print("结果:", "OK" if ok2 else "FAIL")

    # 校准时钟：50% + 负费，模拟三种抓屏条件
    from waa.core.cost_calib import LiveCalibrator  # noqa: E402
    _, states3 = build(rate=50.0, n=420, switch_on=0, switch_off=10 ** 9)

    def clock_case(name, per_logic, m, skip_every=0):
        cal = LiveCalibrator(timeout_s=999.0)
        t = 0.0
        # 一个逻辑帧的真实时长：1X = 1/30s，2X = 1/60s
        dt = 1.0 / (30.0 * m)
        for i, (w, cap, neg) in enumerate(states3):
            if skip_every and i % skip_every == 0:
                t += dt
                continue
            for _ in range(per_logic):
                cal.feed(w, cap, neg, m=m, bullet=False,
                         running=True, t=t)
                t += dt / per_logic
            if cal.done:
                break
        hit = cal.result is not None and abs(cal.result.rate - 50.0) < 0.01
        print(f"校准时钟 {name}: rate={cal.result.rate if cal.result else None} "
              f"frames={cal.observed_frames} {'OK' if hit else 'FAIL'}")
        return hit

    ok3 = clock_case("1X/60fps（每逻辑帧抓两次）", 2, 1.0)
    ok4 = clock_case("2X/60fps（每逻辑帧抓一次）", 1, 2.0)
    ok5 = clock_case("1X/60fps + 每 7 帧漏一帧", 2, 1.0, skip_every=7)
    print("结果:", "OK" if (ok3 and ok4 and ok5) else "FAIL")
    return 0 if (ok and ok2 and ok3 and ok4 and ok5) else 1


if __name__ == "__main__":
    sys.exit(main())
