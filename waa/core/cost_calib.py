"""回费率自动校准：用量到的费用条宽度反推“战前回费率”。

和 CostBarRuler 的思路一致，但只保留一个标量：
    战前回费率 rate% -> 一费帧数 3000/rate -> required = round(frames/30*2^24)
    局内负费不改 required，只把每帧增量 speed 减半（SCALE//60）
所以同样的 60 帧/费，可能是“50% 战前”，也可能是“100% 战前 + 局内负费”；
减号识别（cost_sign）把这两条路分开。

校准用“一个完整周期的不同宽度集合”做匹配：
  - 候选 rate 从 1% 枚举到 115%
  - 按当前的负费状态选 speed，合成该候选一个周期的宽度集合
  - 实测集合里的每个宽度都必须能在候选集合里找到（±1px）
  - 在全部命中的候选里，按“候选多出来、实测没覆盖的宽度数”从少到多排
这个集合匹配不怕漏采：漏了几帧只会让实测集合变小，不会凭空造出错误宽度。
"""

from __future__ import annotations

import time
from dataclasses import dataclass

from .cost_model import (
    EPS, SCALE, SPEED, measured_width, raw_width,
)

MIN_RATE = 1
MAX_RATE = 115
MAX_MODEL_FRAMES = 4000


@dataclass(frozen=True)
class RateCandidate:
    rate: float
    matched: int
    observed: int
    extra: int
    cycle_frames: int

    @property
    def ok(self) -> bool:
        return self.observed > 0 and self.matched == self.observed


def model_cycle_widths(rate_percent: float, negative: bool,
                       max_frames: int = MAX_MODEL_FRAMES) -> list[int]:
    """合成一个回费周期的实测宽度（从相位 0 开始，不含回费那一帧）。"""
    if rate_percent <= 0:
        return []
    required = SCALE
    normal = max(1, int(round(SPEED * rate_percent / 100.0)))
    speed = max(1, normal // 2) if negative else normal
    eps = EPS
    out: list[int] = []
    acc = 0
    for _ in range(max_frames):
        out.append(measured_width(raw_width(acc / float(required))))
        acc += speed
        if acc + eps >= required:
            break
    return out


def _set_of(widths) -> list[int]:
    return sorted({int(w) for w in widths if int(w) > 0 and int(w) < 178})


def deduped_frame_count(widths, caps) -> int:
    """去掉连续重复画面后的帧数（同一逻辑帧被采了两次的那种）。

    0 宽要连条帽明暗一起比；填充只比宽度。慢速档连续多帧同宽也会被并掉，
    所以这个数只会偏小，不会偏大。
    """
    n = 0
    prev = None
    for w, c in zip(widths, caps):
        w = int(w)
        key = (w, bool(c)) if w == 0 else (w, None)
        if key != prev:
            n += 1
            prev = key
    return n


def infer_rate(observed_widths, negative: bool, rates=None,
               observed_frames: int | None = None) -> list[RateCandidate]:
    """返回候选（先全部命中，再按周期帧数、多出来的宽度数排序）。

    observed_frames：这一个周期去重后的帧数。给了它就优先选周期长度最接近的
    候选——同样的宽度集合，81% 和 100% 都可能“全命中”，只有周期长度分得开。
    """
    obs = _set_of(observed_widths)
    if not obs:
        return []
    if rates is None:
        rates = range(MIN_RATE, MAX_RATE + 1)
    cands: list[RateCandidate] = []
    for rate in rates:
        model_raw = model_cycle_widths(float(rate), negative)
        model = _set_of(model_raw)
        if not model:
            continue
        matched = sum(1 for w in obs if any(abs(w - m) <= 1 for m in model))
        extra = sum(1 for m in model if not any(abs(m - w) <= 1 for w in obs))
        cands.append(RateCandidate(
            rate=float(rate), matched=matched, observed=len(obs),
            extra=extra, cycle_frames=len(model_raw),
        ))

    def key(c: RateCandidate):
        diff = (0 if observed_frames is None
                else abs(c.cycle_frames - observed_frames))
        return (diff, c.extra, c.rate)

    ok = [c for c in cands if c.ok]
    if ok:
        ok.sort(key=key)
        return ok
    cands.sort(key=lambda c: (-c.matched, key(c)))
    return cands[:10]


class LiveCalibrator:
    """实时校准：喂逐拍 (width, cap, neg)，等一个完整周期后出结果。

    游戏要处于运行状态、费用条看得见；暂停时宽度不动，等不到回费。
    """

    def __init__(self, timeout_s: float = 60.0):
        self.samples: list[tuple[int, bool, bool | None]] = []
        self.clocks: list[float] = []       # 每拍对应的“已走过多少逻辑帧”
        self.clock = 0.0
        self._last_t: float | None = None
        self._last_running = True
        self._last_m = 1.0
        self._last_bullet = False
        self.t0 = time.perf_counter()
        self.timeout_s = timeout_s
        self.done = False
        self.ok = False
        self.note = "校准中…"
        self.result: RateCandidate | None = None
        self.candidates: list[RateCandidate] = []
        self.negative: bool | None = None
        self.observed_frames = 0
        self.cycles_seen = 0

    def feed(self, width: int, cap_bright: bool, neg: bool | None,
             m: float = 1.0, bullet: bool = False,
             running: bool = True, t: float | None = None) -> None:
        """喂一拍。周期长度按墙钟 × 逻辑帧率算，不数截图。

        1X = 30 逻辑帧/秒，2X = 60，子弹时间 = 6；暂停的区间不计。
        这样慢速档开头那几帧暗零被去重吃掉、或者抓屏漏了几帧，都不会把
        “120 帧/费”算成 117 帧（那会把 50% 校准成 51%）。
        """
        if self.done:
            return
        now = time.perf_counter() if t is None else float(t)
        if self._last_t is not None:
            dt = now - self._last_t
            if 0.0 <= dt <= 0.25 and self._last_running:
                rate = 6.0 if self._last_bullet else 30.0 * self._last_m
                self.clock += dt * rate
        self._last_t = now
        self._last_running = True if running is None else bool(running)
        self._last_m = float(m or 1.0)
        self._last_bullet = bool(bullet)

        self.samples.append((int(width), bool(cap_bright), neg))
        self.clocks.append(self.clock)
        if now - self.t0 > self.timeout_s:
            self._finish(False, "校准超时：没等到完整回费周期（确认游戏在跑、"
                                "费用条可见、没有暂停）")
            return

        widths = [s[0] for s in self.samples]
        caps = [s[1] for s in self.samples]
        negs = [s[2] for s in self.samples]
        cycles = split_cycles(widths, caps, negs)
        self.cycles_seen = len(cycles)
        if not cycles:
            self.note = f"校准中… 等下一次回费（已采 {len(self.samples)} 帧）"
            return
        if len(cycles) < 2:
            # 第一个闭合周期可能是“请求时已经走到一半”的残周期，长度不可信；
            # 等第二个闭合周期再用（它一定是从一次回费走到下一次回费）。
            self.note = "校准中… 再等一个完整周期"
            return

        seg, cneg, seg_caps, start, end = cycles[-1]
        if cneg is None:
            self.note = "校准中… 这个周期负费状态不稳定，继续等"
            return
        clock_frames = self.clocks[end] - self.clocks[start]
        use_clock = clock_frames >= 1.0
        frames = (int(round(clock_frames)) if use_clock
                  else deduped_frame_count(seg, seg_caps))
        if frames < 10:
            self.note = "校准中… 周期太短或时钟没走起来，继续等"
            return
        cands = infer_rate(seg, cneg, observed_frames=frames)
        if not cands or not cands[0].ok:
            top = cands[0] if cands else None
            self.note = ("校准中… 这个周期还没匹配（"
                         + (f"{top.matched}/{top.observed}" if top else "无候选")
                         + "）")
            return

        self.result = cands[0]
        self.candidates = cands
        self.negative = cneg
        self.observed_frames = frames
        self._finish(
            True,
            f"校准完成：战前 {cands[0].rate:g}%"
            f"{' + 局内负费' if cneg else ''}，"
            f"一费 {cands[0].cycle_frames} 帧"
            f"（实测{'时钟' if use_clock else '去重'} {frames} 帧、"
            f"{len(seg)} 张截图）",
        )

    def _finish(self, ok: bool, note: str) -> None:
        self.ok = ok
        self.done = True
        self.note = note


def split_cycles(widths, caps=None, negative_flags=None,
                 min_peak: int = 150, reset_max: int = 30,
                 min_cycle_len: int = 10):
    """按“宽度从峰值大幅掉下来”切周期。

    返回 [(该周期的宽度列表, negative, 该周期的条帽列表, start, end)]；
    negative 为 True/False，中途翻转或全读不到时为 None。只收被下一次回费
    闭合的完整周期；start/end 是两次回费在输入里的下标（end 不含）。
    """
    out = []
    start = 0
    prev = None
    for i, w in enumerate(widths):
        w = int(w)
        if prev is not None and prev >= min_peak and w <= reset_max:
            if i - start >= min_cycle_len:
                seg = [int(x) for x in widths[start:i]]
                seg_caps = ([bool(caps[j]) for j in range(start, i)]
                            if caps is not None else None)
                neg = None
                if negative_flags is not None:
                    vals = [negative_flags[j] for j in range(start, i)
                            if negative_flags[j] is not None]
                    if vals:
                        neg = True if all(vals) else (
                            False if not any(vals) else None)
                out.append((seg, neg, seg_caps, start, i))
            start = i
        prev = w
    return out
