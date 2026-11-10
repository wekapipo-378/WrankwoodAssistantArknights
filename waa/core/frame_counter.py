"""WAA 自己的计数器：费用条 + fp24 增长模型（真实 1920×1080 截图实测）。

坐标（客户区 1920×1080，UI 100%）：
  参考  x1728..1732（半透明底板；两态相同）
  条帽  x1737..1738（紧贴条槽左缘，1/30 会亮起来）
  边框  x1739
  条槽  x1740..1919（180px）
  条体  y811..816（6 行）

宽度不是“每费固定 30 帧”的动画，而是累加器相位的线性函数（见 cost_model）：
  相位 0..1/30   ：宽度 0；一段 0 宽里的最后一帧条帽亮（1/30）
  2/30..29/30    ：宽度 5/12/…/174
  周期最后一帧   ：31 帧周期里相位会到 ~1.0，整条 180（不是另一个 0/30）
回费帧本身（累加器刚减掉 required）相位是 0，画出来是空的。
"""

from __future__ import annotations

import csv
import os
import time
from dataclasses import dataclass

import cv2
import numpy as np

from .cost_calib import infer_rate, split_cycles
from ..paths import bundle_root
from .cost_model import (
    EPS, SCALE, SPEED, measured_width, raw_width,
)
from .cost_sign import read_sign

BAR_X0 = 1740
BAR_X1 = 1920
BAR_Y0 = 811
BAR_Y1 = 817

REF_X0, REF_X1 = 1728, 1733
CAP_X0, CAP_X1 = 1737, 1739

# 2/30..29/30 的白条宽度（真实截图逐帧实测）
FILL_WIDTHS = [
    5, 12, 18, 24, 30, 37, 43, 49, 55, 62,
    68, 74, 80, 86, 93, 99, 105, 111, 118, 124,
    130, 136, 143, 149, 155, 161, 168, 174,
]

FULL_MIN = 178        # 180px：0/30（10 秒后）/ 30/30
CAP_DARK = 5.0        # 条帽-参考 < 5 → 0/30（前 10 秒的纯黑态）

STANDARD_FRAME = (1920, 1080)


def normalize_frame(frame: np.ndarray) -> np.ndarray:
    """16:9 其他分辨率 -> 1920×1080，固定像素 ROI 才能继续用。

    只在“客户区本身就是 16:9 游戏画面”时成立；UI 缩放必须 100%。
    """
    if frame is None:
        return frame
    h, w = frame.shape[:2]
    if w == STANDARD_FRAME[0] and h == STANDARD_FRAME[1]:
        return frame
    interp = cv2.INTER_AREA if w > STANDARD_FRAME[0] else cv2.INTER_CUBIC
    return cv2.resize(frame, STANDARD_FRAME, interpolation=interp)


@dataclass
class BarVisual:
    cap_diff: float
    width: int
    phase: int | None   # 当前周期内的第几帧（0 起）
    kind: str           # "0/30" / "1/30" / "fill" / "full"
    cycle_len: int = 30  # 当前周期总帧数（100%=30，50%=60……）
    uncertainty: int = 0  # 条连续多帧同宽时的可能帧数上限（实际帧 ∈ [frame, frame+uncertainty]）
    neg: bool = False    # 局内负费（数字左边减号），speed 减半
    # 执行器要读的字段（倍速/子弹时间先给默认值，后面接 WAA 的按钮识别）
    t: float = 0.0
    m: float = 1.0
    bullet: bool = False
    running: bool = True
    pause_drawn: bool = False
    speed_drawn: bool = False
    mode: str = "bar"


def _run_from_slot_start(body: np.ndarray, threshold: float,
                         min_rows: int = 3) -> int:
    """填充是从条槽左端开始的一段连续亮区；只量这一段，避免背景亮点拉长跨度。

    第一列允许被抗锯齿压暗 1~2 列，起点在左端 3 列内找。
    """
    sub = body[:, BAR_X0:BAR_X1]
    if sub.size == 0:
        return 0
    cols = (sub > threshold).sum(axis=0) >= min_rows
    n = len(cols)
    start = None
    for i in range(min(3, n)):
        if cols[i]:
            start = i
            break
    if start is None:
        return 0
    end = start
    while end + 1 < n and cols[end + 1]:
        end += 1
    return end + 1  # 宽度从条槽左缘 x1740 算起


def _track_level(body: np.ndarray) -> float:
    """空槽亮度：最右 25 列列均值的**最小值**。

    填充是亮白、空槽是暗色叠加；宽度 <=174 时右端必定还有空槽，最小值就是它。
    （整条 180 的情况在调用前已经判掉。）
    """
    cm = body[:, BAR_X0:BAR_X1].mean(axis=0)
    if cm.size == 0:
        return 0.0
    return float(np.min(cm[-25:]))


def classify_frame(gray: np.ndarray) -> BarVisual:
    """输入完整客户区灰度图（1920×1080），背景自适应。"""
    body = gray[BAR_Y0:BAR_Y1, :]
    ref = float(body[:, REF_X0:REF_X1].mean())
    cap = float(body[:, CAP_X0:CAP_X1].mean())
    cap_diff = cap - ref

    # 1) 先判整条纯白：填充是饱和白 255；空槽即使背景再亮也到不了连续 178 列 >200
    full_width = _run_from_slot_start(body, 200)
    if full_width >= FULL_MIN:
        return BarVisual(cap_diff, full_width, 30, "full")

    # 2) 普通填充：阈值跟着空槽亮度走（背景亮则阈值抬高），下限 200
    track = _track_level(body)
    threshold = max(200.0, track + 60.0)
    width = _run_from_slot_start(body, threshold)

    if width >= 4:
        idx = min(range(len(FILL_WIDTHS)), key=lambda i: abs(FILL_WIDTHS[i] - width))
        return BarVisual(cap_diff, width, idx + 2, "fill")
    if cap_diff < CAP_DARK:
        return BarVisual(cap_diff, width, 0, "0/30")
    return BarVisual(cap_diff, width, 1, "1/30")


_START_TEMPLATES = None
_START_TEMPLATE_TRIED = False
START_ROI = (1400, 500, 1900, 1000)
START_SCALE = 0.5


def _get_start_templates():
    global _START_TEMPLATES, _START_TEMPLATE_TRIED
    if not _START_TEMPLATE_TRIED:
        _START_TEMPLATE_TRIED = True
        path = os.path.join(bundle_root(), "waa", "resources",
                            "StartOperation.png")
        tpl = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
        if tpl is None:
            _START_TEMPLATES = []
        else:
            _START_TEMPLATES = [
                cv2.resize(tpl, None, fx=START_SCALE * s,
                           fy=START_SCALE * s, interpolation=cv2.INTER_AREA)
                for s in (1.0, 0.9)
            ]
    return _START_TEMPLATES


def start_screen_score(gray: np.ndarray) -> float:
    """编队界面“开始行动”按钮的模板分数。

    编队界面的右上角“快捷编队”也可能被按钮识别误看成暂停图标，
    所以战斗判断里要先把编队界面排除掉。ROI 和模板都缩到一半再匹配，
    每帧约 1.5ms；全分辨率要 6ms，60 帧下会白吃 CPU。
    """
    templates = _get_start_templates()
    if not templates or gray is None:
        return -1.0
    x0, y0, x1, y1 = START_ROI
    roi = gray[y0:y1, x0:x1]
    if roi.size == 0:
        return -1.0
    roi = cv2.resize(roi, None, fx=START_SCALE, fy=START_SCALE,
                     interpolation=cv2.INTER_AREA)
    best = -1.0
    for t in templates:
        if t.shape[0] > roi.shape[0] or t.shape[1] > roi.shape[1]:
            continue
        best = max(best, float(cv2.matchTemplate(
            roi, t, cv2.TM_CCOEFF_NORMED).max()))
    return best


def looks_like_start_screen(gray: np.ndarray, threshold: float = 0.55) -> bool:
    return start_screen_score(gray) >= threshold


def looks_like_cost_bar(gray: np.ndarray) -> bool:
    """这一帧像不像战斗里的费用条（用来等开局/防止在别界面起锚）。"""
    body = gray[BAR_Y0:BAR_Y1, :]
    if body.shape[0] < 6:
        return False
    v = classify_frame(gray)
    if v.kind == "full":
        return True
    # 战斗里的空槽是 40~140 的深灰；加载/编队界面的底图通常更暗或更亮
    cm = body[:, BAR_X0:BAR_X1].mean(axis=0)
    track = float(np.min(cm[-25:])) if cm.size else 0.0
    if not (40.0 <= track <= 140.0):
        return False
    right = body[:, BAR_X1 - 6:BAR_X1]
    return (25.0 <= float(right.mean()) <= 160.0
            and float(right.std()) < 25.0)


def measure_bar(gray: np.ndarray):
    """返回 (width, cap_bright, cap_diff)。

    宽度用“从槽左端开始的连续亮段”，阈值跟着空槽亮度走（背景自适应），
    和 fp24 模型校准过（模型宽 - 1）。
    """
    body = gray[BAR_Y0:BAR_Y1, :]
    ref = float(body[:, REF_X0:REF_X1].mean())
    cap = float(body[:, CAP_X0:CAP_X1].mean())
    cap_diff = cap - ref
    if body.size == 0:
        return 0, False, cap_diff
    full = _run_from_slot_start(body, 200.0)
    if full >= FULL_MIN:
        return full, cap_diff >= CAP_DARK, cap_diff
    track = _track_level(body)
    threshold = max(200.0, track + 60.0)
    width = _run_from_slot_start(body, threshold)
    return width, cap_diff >= CAP_DARK, cap_diff


class WAAFrameCounter:
    """费用条 + fp24 累加器：视觉宽度 -> 全局逻辑帧（0 基）。

    每个逻辑帧 acc += speed；acc + eps >= required 时回一费并 acc -= required。
    宽度 = round(1.03892×180×phase − 7.0056) + 1（再转成实测量法，见 cost_model）。
    每次采样从当前累加器向前模拟 0..MAX_JUMP 帧，取第一个对得上的状态：
      - 状态没变 = 同一游戏帧的重复采样，不涨帧（暂停时正好冻结）
      - 向前 k 帧 = 中间漏采了 k 帧，补 k
      - 都对不上 = 误判，冻结这一帧
    局内负费（费用数字左边有减号）时 speed 从 SCALE//30 切成 SCALE//60，
    累加器/相位原样保留，不重新起锚。
    """

    MAX_JUMP = 12
    MATCH_TOL = 1

    def __init__(self, rate_percent: float = 100.0,
                 max_frames: int = 200000,
                 neg_override: bool | None = None,
                 neg_speed: int | None = None,
                 neg_declared: bool = True,
                 rate_candidates: list[float] | None = None,
                 auto_detect: bool = True):
        self.rate_percent = float(rate_percent)
        self.neg_override = neg_override
        self.neg_declared = bool(neg_declared)
        self.spec_error = ""
        # 自动检测只在轴里声明的候选档位之间切换；负费用候选 // 2。
        self.rate_candidates = self._normalize_candidates(
            rate_candidates, self.rate_percent)
        self.auto_detect = bool(auto_detect and self.rate_candidates)
        self._pending_rate: float | None = None
        self._pending_hits = 0
        # 周期级检测：攒一个完整回费周期的条宽，反推当前档位。
        self._rate_samples: list[tuple[int, bool, bool | None]] = []
        self._cycle_pending_rate: float | None = None
        self._cycle_pending_hits = 0
        # 冷却范围固定为 1（2^24）；回费速度只改每帧 speed。
        self.required = SCALE
        self.eps = EPS
        self.normal_speed = (
            max(1, int(round(SPEED * self.rate_percent / 100.0)))
            if self.rate_percent > 0 else 0)
        self.neg_speed = (int(neg_speed) if neg_speed
                          else max(1, self.normal_speed // 2))
        self.speed = self.normal_speed
        self.neg = False
        self.acc = 0
        self.since = 0
        self.cycles = 0
        self.frame = 0
        self.uncertainty = 0
        self.last: BarVisual | None = None
        self.warn = ""
        self.csv_path: str | None = None
        self.last_delta = 0
        self.manual_rate: float | None = None  # 当前手动回费速度；None=默认/自动
        self._need_anchor = True
        self._t0 = time.perf_counter()
        self._n = 0
        self._csv_file = None
        self._csv_writer = None

    @staticmethod
    def _normalize_candidates(values, fallback: float) -> list[float]:
        if values is None:
            values = []
        elif not isinstance(values, (list, tuple, set)):
            values = [values]
        out = []
        seen = set()
        for v in list(values) + [fallback]:
            try:
                rate = float(v)
            except (TypeError, ValueError):
                continue
            if rate <= 0:
                continue
            rate = round(rate, 4)
            if rate in seen:
                continue
            seen.add(rate)
            out.append(rate)
        return sorted(out)

    # ---- 对外接口 ----
    def reset(self, frame: int = 0) -> None:
        self.frame = int(frame)
        self.acc = 0
        self.since = 0
        self.cycles = 0
        self.uncertainty = 0
        # 新的一局/重新起锚从正常回费开始；负费由下一帧画面重新判断。
        self.neg = False
        self.speed = self.normal_speed
        self.spec_error = ""
        self._pending_rate = None
        self._pending_hits = 0
        self._rate_samples.clear()
        self._cycle_pending_rate = None
        self._cycle_pending_hits = 0
        self._need_anchor = True
        self.last = None
        self.warn = ""
        self._t0 = time.perf_counter()

    def resync(self) -> None:
        """下一个视觉状态重新对齐，帧号保持当前值。"""
        self._need_anchor = True

    def advance_blind(self, frames: int = 1) -> int:
        """盲步进：费用条静止时，按当前 speed 把模型推进 frames 帧。

        只在执行器确认游戏停住、子弹时间仍在、且目标落在当前
        uncertainty 区间内时调用。它不产生新的视觉信息，只是把
        “输入已走了一帧”记进模型。
        """
        n = max(0, int(frames))
        for _ in range(n):
            self.acc, self.since, _ = self._advance(
                self.acc, self.since, self.speed)
            self.frame += 1
        if n:
            self.uncertainty = max(0, self.uncertainty - n)
            self.last_delta = n
            self.warn = ""
        return self.frame

    def set_initial_frame(self, frame: int) -> None:
        """中途启动/调试：强制当前全局帧。"""
        self.reset(max(0, int(frame)))

    def current_frame(self) -> int:
        return self.frame

    # ---- fp24 ----
    def _width_at(self, acc: int) -> int:
        if self.required <= 0:
            return 0
        return measured_width(raw_width(acc / float(self.required)))

    def _cap_at(self, acc: int, speed: int) -> bool:
        """当前帧宽 0 时条帽亮不亮 = 下一帧是否开始画出填充。"""
        nxt = acc + speed
        if nxt + self.eps >= self.required:
            nxt -= self.required
        return self._width_at(nxt) > 0

    def _error(self, acc: int, width: int, cap_bright: bool,
               speed: int) -> int:
        """这个状态和实测画面的像素误差；0 最好。"""
        model = self._width_at(acc)
        obs_full = width >= FULL_MIN
        if model >= FULL_MIN:
            return 0 if obs_full else 40 + abs(model - width)
        if obs_full:
            return 40 + abs(model - width)
        if model == 0:
            if width != 0:
                return 40 + width
            return 0 if self._cap_at(acc, speed) == cap_bright else 30
        return abs(model - width)

    def _state_ok(self, acc: int, width: int, cap_bright: bool,
                  speed: int) -> bool:
        return self._error(acc, width, cap_bright, speed) <= self.MATCH_TOL

    def _advance(self, acc: int, since: int, speed: int):
        acc += speed
        grant = acc + self.eps >= self.required
        if grant:
            acc -= self.required
            since = 0
        else:
            since += 1
        return acc, since, grant

    def _search(self, width: int, cap_bright: bool, limit: int, speed: int,
                return_error: bool = False):
        """从当前状态向前找 0..limit 帧里像素误差最小的状态。

        不能“第一个落在 ±1px 内就收”：慢速档（比如 50% 战前 + 负费 =
        120 帧/费）相邻两帧宽度只差 1px，第一个命中的永远是上一帧，显示会
        一直慢 1 帧。取整段里误差最小的那个才对。

        返回 (k, acc, since, uncertainty)。uncertainty = 与最优误差相同的
        状态里最靠后的那个多走了几帧；条连续多帧同宽时实际帧不确定，
        显示要给出 [frame, frame+uncertainty] 区间。
        return_error=True 时额外返回 best_err，供候选档位比较。
        """
        acc, since = self.acc, self.since
        states = []
        for k in range(limit + 1):
            err = self._error(acc, width, cap_bright, speed)
            states.append((k, acc, since, err))
            if k == limit:
                break
            acc, since, _ = self._advance(acc, since, speed)
        best_err = min(s[3] for s in states)
        if best_err > self.MATCH_TOL:
            return None
        tied = [s for s in states if s[3] == best_err]
        k, acc, since, _ = min(tied, key=lambda s: s[0])
        k_max = max(s[0] for s in tied)
        result = (k, acc, since, k_max - k)
        if return_error:
            return result + (best_err,)
        return result

    def _candidate_speeds(self, negative: bool):
        """把声明的候选回费率展开成 (rate, normal_speed, speed)。"""
        out = []
        for rate in self.rate_candidates:
            normal = max(1, int(round(SPEED * rate / 100.0)))
            speed = max(1, normal // 2) if negative else normal
            out.append((float(rate), normal, speed))
        return out

    def _active_speeds(self, negative: bool):
        """自动检测开时用候选集；手动速度时只用当前这一档。"""
        if self.auto_detect:
            return self._candidate_speeds(negative)
        speed = self.neg_speed if negative else self.normal_speed
        return [(float(self.rate_percent), int(self.normal_speed), int(speed))]

    def _apply_rate(self, rate: float, negative: bool) -> None:
        """切换到某个声明档位；保留 acc/since 的相位。"""
        old = float(self.rate_percent)
        rate = float(rate)
        normal = max(1, int(round(SPEED * rate / 100.0)))
        self.rate_percent = rate
        self.normal_speed = normal
        self.neg_speed = max(1, normal // 2)
        self.neg = bool(negative)
        self.speed = self.neg_speed if self.neg else self.normal_speed
        self._pending_rate = None
        self._pending_hits = 0
        self._rate_samples.clear()
        self._cycle_pending_rate = None
        self._cycle_pending_hits = 0
        if abs(old - rate) > 1e-9:
            print(f"[rate] 自动切换 {old:g}% -> {rate:g}%"
                  f"{'（负费）' if negative else ''}")

    def _best_candidate(self, width: int, cap_bright: bool, negative: bool):
        """当前档对不上时，在候选档里找误差最小的一个。"""
        best = None
        for rate, normal, speed in self._candidate_speeds(negative):
            if speed == self.speed:
                continue
            found = self._search(width, cap_bright, self.MAX_JUMP, speed,
                                 return_error=True)
            if found is None:
                continue
            err = found[4]
            key = (err, abs(rate - self.rate_percent), rate)
            if best is None or key < best[0]:
                best = (key, rate, normal, speed, found)
        if best is None:
            return None
        return best[1], best[2], best[3], best[4]

    def _confirm_candidate(self, rate: float) -> bool:
        """新档位连续命中若干拍才切换，避免单帧抖动。"""
        if (self._pending_rate is not None
                and abs(self._pending_rate - rate) < 1e-9):
            self._pending_hits += 1
        else:
            self._pending_rate = float(rate)
            self._pending_hits = 1
        return self._pending_hits >= 2

    def _feed_rate_samples(self, width: int, cap_bright: bool,
                           neg_obs: bool | None) -> None:
        """用完整回费周期反推档位：21% 这类检测值会被吸附到 20%。

        逐帧候选匹配分不清 100% 和 110%（12 帧窗口内条宽太接近），
        但整个周期的宽度集合可以。只在候选集大于一档时启用。
        """
        if not self.auto_detect or len(self.rate_candidates) <= 1:
            self._rate_samples.clear()
            return
        self._rate_samples.append((int(width), bool(cap_bright), neg_obs))
        if len(self._rate_samples) > 2000:
            del self._rate_samples[:-1200]
        widths = [s[0] for s in self._rate_samples]
        caps = [s[1] for s in self._rate_samples]
        negs = [s[2] for s in self._rate_samples]
        cycles = split_cycles(widths, caps, negs)
        if not cycles:
            return
        seg, cneg, _seg_caps, _start, _end = cycles[-1]
        # 只接受从回费后的低宽度开始的完整周期。缓冲区如果是在中途
        # 被清空的，第一段会从任意宽度开始，这种残周期不能用来判定档位。
        if cneg is None or len(seg) < 10 or int(seg[0]) > 12:
            return
        cands = infer_rate(seg, cneg, rates=self.rate_candidates)
        if not cands or not cands[0].ok or cands[0].extra > 0:
            return
        rate = float(cands[0].rate)
        if abs(rate - self.rate_percent) < 1e-9:
            self._cycle_pending_rate = None
            self._cycle_pending_hits = 0
            return
        if (self._cycle_pending_rate is not None
                and abs(self._cycle_pending_rate - rate) < 1e-9):
            self._cycle_pending_hits += 1
        else:
            self._cycle_pending_rate = rate
            self._cycle_pending_hits = 1
        if self._cycle_pending_hits >= 2:
            self._apply_rate(rate, cneg)

    def _reset_pending(self) -> None:
        self._pending_rate = None
        self._pending_hits = 0

    def _period(self, speed: int) -> int:
        if self.required <= 0:
            return 0
        acc = 0
        for n in range(1, 10001):
            acc += speed
            if acc + self.eps >= self.required:
                return n
        return 10000

    def _frames_until_grant(self, speed: int) -> int:
        if self.required <= 0:
            return 0
        acc, since = self.acc, self.since
        for n in range(1, 10001):
            acc, since, grant = self._advance(acc, since, speed)
            if grant:
                return n
        return 0

    def _set_visual(self, v: BarVisual, since: int, speed: int) -> None:
        v.phase = since
        v.cycle_len = since + self._frames_until_grant(speed)
        v.neg = self.neg
        v.uncertainty = self.uncertainty

    def update(self, gray: np.ndarray) -> BarVisual:
        width, cap_bright, cap_diff = measure_bar(gray)
        if width >= FULL_MIN:
            kind = "full"
        elif width > 0:
            kind = "fill"
        else:
            kind = "1/30" if cap_bright else "0/30"
        v = BarVisual(cap_diff, width, None, kind)
        v.t = time.perf_counter() - self._t0
        self._n += 1

        if self.required <= 0:
            self.warn = "回费速度为 0：费用条不增长，帧号不可用；请检查回费速度设置"
            self.last_delta = 0
            self._log(v, 0)
            self.last = v
            return v

        neg_obs = (self.neg_override if self.neg_override is not None
                   else read_sign(gray))
        # 负费固定按当前速度 ÷ 2；不再要求轴里额外声明“负费档”。
        self._feed_rate_samples(width, cap_bright, neg_obs)

        if self._need_anchor:
            neg_state = self.neg if neg_obs is None else bool(neg_obs)
            best = None
            for rate, _normal, speed in self._active_speeds(neg_state):
                period = max(1, self._period(speed))
                found = self._search(width, cap_bright, period, speed,
                                     return_error=True)
                if found is None and width >= FULL_MIN:
                    # 满条只出现在“这一帧还没回费、相位贴在 1.0 上”的那一拍。
                    # 从 acc=0 推进的普通周期里这个状态会被回费扣掉，补两个候选。
                    for start_acc in (self.required - 1,
                                      self.required - speed // 2):
                        if self._state_ok(start_acc, width, cap_bright, speed):
                            found = (0, start_acc, max(0, period - 1), 0, 0)
                            break
                if found is None:
                    continue
                err = found[4]
                key = (err, abs(rate - self.rate_percent), rate)
                if best is None or key < best[0]:
                    best = (key, rate, speed, found)
            if best is None:
                self.warn = (
                    f"费用条起锚失败（width={width}）：帧号不可用；"
                    "请检查费用条是否被遮挡、游戏 UI 是否为 100%")
                self.last_delta = 0
                self.uncertainty = 0
                self._log(v, 0)
                self.last = v
                return v
            _, rate, speed, found = best
            self._apply_rate(rate, neg_state)
            _, acc, since, unc, _ = found
            self.acc = acc
            self.since = since
            self.uncertainty = unc
            self._need_anchor = False
            self.warn = ""
            self.last_delta = 0
            self._set_visual(v, since, speed)
            self._log(v, 0)
            self.last = v
            return v

        # 负号和速度在游戏里同步生效，但这一拍看到的条宽是**上一帧**用
        # 上一帧 speed 画出来的；所以先用旧速度匹配，减号刚翻转时用新
        # 速度做兜底。这不是“负费检测延迟”，只是画面更新顺序。
        use_neg = self.neg if neg_obs is None else bool(neg_obs)
        switched = False
        found = self._search(width, cap_bright, self.MAX_JUMP, self.speed,
                             return_error=True)
        if found is None and neg_obs is not None and bool(neg_obs) != self.neg:
            alt_speed = self.neg_speed if neg_obs else self.normal_speed
            if alt_speed != self.speed:
                found = self._search(width, cap_bright, self.MAX_JUMP,
                                     alt_speed, return_error=True)
                if found is not None:
                    self._apply_rate(self.rate_percent, use_neg)
                    switched = True
        if found is None and self.auto_detect:
            best = self._best_candidate(width, cap_bright, use_neg)
            if best is not None:
                rate, _normal, _speed, candidate_found = best
                if self._confirm_candidate(rate):
                    self._apply_rate(rate, use_neg)
                    found = candidate_found
                    switched = True
                else:
                    self.warn = f"检测到回费速度 {rate:g}%，确认中"
                    self.last_delta = 0
                    self.uncertainty = 0
                    self._set_visual(v, self.since, self.speed)
                    self._log(v, 0)
                    self.last = v
                    return v
        if found is None:
            self._reset_pending()
            self.warn = (
                f"费用条匹配失败（width={width}）：帧号暂时冻结；"
                "常见原因是费用条连续 12 帧以上未变化或被遮挡")
            self.last_delta = 0
            self.uncertainty = 0
            self._set_visual(v, self.since, self.speed)
            self._log(v, 0)
            self.last = v
            return v

        k, acc, since, unc, _err = found
        self._reset_pending()
        self.acc = acc
        self.since = since
        self.uncertainty = unc
        self.frame += k
        self.last_delta = k
        self.warn = ""
        if neg_obs is not None:
            self.neg = use_neg
            if not switched:
                self.speed = self.neg_speed if self.neg else self.normal_speed
        self._set_visual(v, since, self.speed)
        self._log(v, k)
        self.last = v
        return v

    def _log(self, v: BarVisual, delta: int) -> None:
        if not self.csv_path:
            return
        if self._csv_file is None:
            self._csv_file = open(self.csv_path, "w", encoding="utf-8",
                                  newline="", buffering=1)
            self._csv_writer = csv.writer(self._csv_file)
            self._csv_writer.writerow([
                "idx", "t", "frame", "unc", "delta", "phase", "width",
                "cap_diff", "kind", "neg", "warn",
            ])
        self._csv_writer.writerow([
            self._n, f"{v.t:.4f}", self.frame, v.uncertainty, delta,
            v.phase, v.width, f"{v.cap_diff:.1f}", v.kind, int(v.neg),
            self.warn,
        ])

    def close_log(self) -> None:
        if self._csv_file is not None:
            self._csv_file.close()
            self._csv_file = None
            self._csv_writer = None
