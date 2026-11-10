"""离线测试：按轴设置候选档位自动检测回费速度。"""

from __future__ import annotations

import os
import sys
from unittest.mock import patch

import numpy as np

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa.core.cost_model import SPEED, build_sequence  # noqa: E402
from waa.core.frame_counter import WAAFrameCounter  # noqa: E402
from waa.core.waa_counter import WAACounter  # noqa: E402


def check(cond, label):
    print(("PASS " if cond else "FAIL ") + label)
    return bool(cond)


def cycle_samples(rate):
    """一个完整周期的 (width, cap)，cap 是真实条帽状态。"""
    widths, _counts, caps, _cs, cycle_len = build_sequence(
        rate, max_frames=5000)
    return [(int(widths[i]), bool(caps[i])) for i in range(int(cycle_len[0]))]


def find_mismatch(counter, true_speed, limit=80):
    """找一个“当前 100% 对不上、110% 对得上”的真实帧。"""
    acc = since = 0
    for n in range(limit):
        width = counter._width_at(acc)
        cap = counter._cap_at(acc, true_speed)
        counter.acc, counter.since, counter.frame = acc, since, n
        cur = counter._search(width, cap, counter.MAX_JUMP,
                              counter.normal_speed)
        cand = counter._search(width, cap, counter.MAX_JUMP, true_speed)
        if cur is None and cand is not None:
            return n, acc, since, width, cap
        acc, since, _ = counter._advance(acc, since, true_speed)
    return None


def main() -> int:
    checks = []
    counter = WAAFrameCounter(100, rate_candidates=[100, 110, 50])
    true_speed = max(1, int(round(SPEED * 110 / 100.0)))

    # 1) 候选档位只来自设置。
    checks.append(check(counter.rate_candidates == [50.0, 100.0, 110.0],
                        f"候选 = {counter.rate_candidates}"))
    speeds = dict((r, s) for r, _n, s in counter._candidate_speeds(False))
    checks.append(check(speeds[110.0] == true_speed,
                        f"110% speed = {speeds[110.0]}"))

    # 2) 逐帧候选：100% -> 50% 这种大变化能直接选中。
    c50 = WAAFrameCounter(100, rate_candidates=[100, 50])
    true50 = max(1, int(round(SPEED * 50 / 100.0)))
    hit = find_mismatch(c50, true50, limit=120)
    checks.append(check(hit is not None, "找到 100/50 可区分的帧"))
    if hit is not None:
        n, acc, since, width, cap = hit
        c50.acc, c50.since, c50.frame = acc, since, n
        best = c50._best_candidate(width, cap, False)
        checks.append(check(best is not None and best[0] == 50.0,
                            f"选中的候选 = {best[0] if best else None}"))
        checks.append(check(not c50._confirm_candidate(50.0),
                            "第一次命中只确认中"))
        checks.append(check(c50._confirm_candidate(50.0),
                            "第二次命中才切换"))

    # 3) 周期级：100% -> 110% 这种接近档位靠完整周期识别。
    samples110 = cycle_samples(110)
    counter = WAAFrameCounter(100, rate_candidates=[100, 110, 50])
    for w, cap in samples110:
        counter._feed_rate_samples(w, cap, False)
    counter._feed_rate_samples(*samples110[0], False)
    checks.append(check(counter.rate_percent == 100.0,
                        "第一周期只确认中，不切换"))
    counter._feed_rate_samples(*samples110[1], False)
    checks.append(check(counter.rate_percent == 110.0,
                        f"第二周期切到 {counter.rate_percent:g}%"))
    checks.append(check(counter.speed == true_speed,
                        f"切换后 speed={counter.speed}"))
    checks.append(check(counter.neg is False, "正费状态"))

    # 4) 周期级：本来就是 100% 时不要乱切。
    samples100 = cycle_samples(100)
    c100 = WAAFrameCounter(100, rate_candidates=[100, 110, 50])
    for w, cap in samples100:
        c100._feed_rate_samples(w, cap, False)
    c100._feed_rate_samples(*samples100[0], False)
    c100._feed_rate_samples(*samples100[1], False)
    checks.append(check(c100.rate_percent == 100.0,
                        f"100% 周期保持 {c100.rate_percent:g}%"))

    # 5) 真正走 update：喂完整 110% 周期，自动切档。
    cu = WAAFrameCounter(100, rate_candidates=[100, 110, 50])
    cu._need_anchor = False
    seq = samples110 + [samples110[0], samples110[1]]
    with patch("waa.core.frame_counter.measure_bar",
               side_effect=[(w, cap, 0.0) for w, cap in seq]), \
            patch("waa.core.frame_counter.read_sign", return_value=False):
        for _ in seq:
            cu.update(np.zeros((10, 10), dtype=np.uint8))
    checks.append(check(cu.rate_percent == 110.0,
                        f"update 自动切到 {cu.rate_percent:g}%"))

    # 6) 负费候选 = 每档 // 2。
    counter._apply_rate(110.0, True)
    checks.append(check(counter.neg is True, "负费切换"))
    checks.append(check(counter.speed == true_speed // 2,
                        f"负费 speed={counter.speed}"))

    # 5) 手动回费速度关闭自动检测，恢复自动后重新打开。
    wc = WAACounter(100, rate_candidates=[100, 110, 50])
    wc._set_rate_override(80.0)
    checks.append(check(wc.counter.auto_detect is False,
                        "手动 80% 关闭自动检测"))
    checks.append(check(wc.counter.rate_percent == 80.0,
                        f"手动档 = {wc.counter.rate_percent:g}%"))
    checks.append(check(wc.counter.neg_speed == wc.counter.normal_speed // 2,
                        "手动负费 = 半速"))
    wc._set_rate_override(None)
    checks.append(check(wc.counter.auto_detect is True,
                        "自动模式恢复检测"))
    checks.append(check(wc.counter.rate_percent == 100.0,
                        f"恢复基准 = {wc.counter.rate_percent:g}%"))

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
