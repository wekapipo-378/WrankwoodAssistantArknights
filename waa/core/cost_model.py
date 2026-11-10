"""fp24 费用计时 + 相位→宽度曲线：支持 0–115% 的回费速率。

模型来自实测（CostBarRulerPlus 等参考实现的同一套结论）：

    required = round(一费帧数 / 30 * 2^24)
    每逻辑帧 accumulator += 559240 (2^24 // 30)
    accumulator + 167 >= required 时回费，accumulator -= required

条宽 = round(1.03892 * 条槽宽 * phase - 隐藏长度) + 1，clamp 到条槽宽；
宽度 <2 时画面上等于没画。

这里把它预生成成一条“每帧长什么样”的序列，计数器只需要在序列里向前匹配。
"""

from __future__ import annotations

import math

import numpy as np

LOGICAL_FPS = 30
SCALE = 1 << 24
SPEED = SCALE // LOGICAL_FPS       # 559240
NEG_SPEED = SCALE // 60            # 279620：局内负费，每帧只攒一半


def epsilon_for(required: int) -> int:
    """把冷却阈值 1e-5 换算成 fp24 整数。

    条件是 cooldown < 1e-5（严格小于），所以取 floor：
    100% 时 required=2^24 -> 167；80% 时 required=2^24/0.8 -> 209。
    """
    return int(1e-5 * required) if required > 0 else 0


EPS = epsilon_for(SCALE)           # 167：100% 基准的阈值
BAR_W = 180
BAR_LENGTH_RATIO = 1.03892
MIN_VISIBLE = 2

# 计数用状态
ZERO = 0
FILL = 1
FULL = 2


def rate_to_frames(percent: float) -> float:
    """回费速率百分比 -> 一费多少逻辑帧。100% = 30 帧。"""
    if percent <= 0:
        return float("inf")
    return LOGICAL_FPS * 100.0 / float(percent)


def rate_from_frames(frames: float) -> float:
    """一费多少逻辑帧 -> 回费速率百分比。30 帧 = 100%，60 帧 = 50%。"""
    if frames is None or frames <= 0:
        return 0.0
    return LOGICAL_FPS * 100.0 / float(frames)


def required_from_frames(frames: float) -> int:
    return int(round(frames / float(LOGICAL_FPS) * SCALE))


def parse_declared_speeds(values) -> list[float]:
    """把“可出现回费速度”的单元格序列解析成百分比列表。

    Excel 里按出现顺序写**等效**速度，例如会欠费：
        第一格 1（=100%），下一格 0.5（=50%）
    也接受 100% / 50% / 100 / 50 的写法。第一格是基准速度，
    后面的值是该局中会切换到的等效速度。
    """
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        values = [values]
    out: list[float] = []
    for v in values:
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        try:
            if s.endswith("%"):
                rate = float(s[:-1])
            else:
                f = float(s)
                # 1 / 0.5 这种是倍率；100 / 50 这种是百分比
                rate = f * 100.0 if f <= 2.0 else f
        except ValueError:
            continue
        if 0.0 < rate <= 1000.0:
            out.append(round(rate, 4))
    return out


def _round_half_away(v: float) -> int:
    if v >= 0.0:
        return int(math.floor(v + 0.5))
    return int(math.ceil(v - 0.5))


def raw_width(phase: float) -> int:
    """相位 0..1 对应的游戏内绘制宽度（含那条 1px 边）。"""
    bar_len = BAR_LENGTH_RATIO * BAR_W
    hidden = bar_len - BAR_W
    p = min(1.0, max(0.0, phase))
    raw = _round_half_away(bar_len * p - hidden) + 1
    if raw < MIN_VISIBLE:
        return 0
    return min(raw, BAR_W)


def measured_width(raw: int) -> int:
    """游戏绘制宽度 -> 我们用阈值量出来的跨度。

    实测：raw=6 量到 5、raw=13 量到 12……整体比 raw 少 1；
    只有整条 clamp 到 180 时量到的也是 180。
    """
    if raw >= BAR_W:
        return BAR_W
    if raw <= 0:
        return 0
    return raw - 1


def build_sequence(rate_percent: float, max_frames: int = 200000,
                   speed: int | None = None):
    """返回 (widths, counts, caps, cycle_start, cycle_len) 五个数组。

    widths[i]       第 i 帧量出来的宽度
    counts[i]       ZERO / FILL / FULL
    caps[i]         宽度为 0 时：0=暗条帽，1=亮条帽（一段 0 宽里的最后一帧亮）
    cycle_start[i]  当前费用周期从第几帧开始（回费后的第一帧）
    cycle_len[i]    当前周期的长度（30、31、60、61……）

    固定 required=2^24；回费速度只改每帧 speed。
    speed 不传时按 rate_percent 算；局内负费传 normal_speed//2。
    """
    if rate_percent <= 0:
        return None
    if speed is None:
        speed = max(1, int(round(SPEED * rate_percent / 100.0)))
    if speed <= 0:
        return None
    required = SCALE
    eps = EPS

    widths = np.zeros(max_frames, np.int16)
    counts = np.zeros(max_frames, np.int8)
    cycle_start = np.zeros(max_frames, np.int32)

    acc = 0
    grant_starts = [0]
    for i in range(max_frames):
        phase = acc / float(required)
        raw = raw_width(phase)
        if raw >= BAR_W:
            widths[i] = BAR_W
            counts[i] = FULL
        elif raw >= MIN_VISIBLE:
            widths[i] = measured_width(raw)
            counts[i] = FILL
        else:
            widths[i] = 0
            counts[i] = ZERO
        if grant_starts:
            cycle_start[i] = grant_starts[-1]

        acc += speed
        if acc + eps >= required:
            acc -= required
            grant_starts.append(i + 1)
            if i + 1 < max_frames:
                cycle_start[i + 1] = i + 1

    caps = np.zeros(max_frames, np.int8)
    for i in range(max_frames):
        if counts[i] == ZERO and (i + 1 >= max_frames or counts[i + 1] != ZERO):
            caps[i] = 1

    cycle_len = np.zeros(max_frames, np.int16)
    for k, start in enumerate(grant_starts):
        if start >= max_frames:
            continue
        end = grant_starts[k + 1] if k + 1 < len(grant_starts) else max_frames
        cycle_len[start:end] = end - start
    return widths, counts, caps, cycle_start, cycle_len
