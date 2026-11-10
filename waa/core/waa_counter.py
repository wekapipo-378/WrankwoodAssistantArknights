"""WAA 计数器：费用条状态机 + 自己的按钮识别。

帧号完全由费用条视觉状态机给出（WAAFrameCounter）；
暂停/倍速/子弹时间由 waa.core.buttons 自己识别，不用外部模板。

每帧写一行合并 CSV（out/waa_run.csv），方便事后用费用条帧率反过来校验：
    idx, t, frame, unc, delta, phase, width, cap_diff, kind,
    running, m, bullet, neg, pause_drawn, speed_drawn, warn
unc = 费用条连续多帧同宽时，实际帧可能比 frame 大 0..unc。
"""

from __future__ import annotations

import csv
import threading
import time

import cv2

from .buttons import read_buttons
from .cost_model import EPS, SCALE, SPEED
from .frame_counter import (WAAFrameCounter, BarVisual, BAR_Y0, BAR_Y1,
                            looks_like_cost_bar, looks_like_start_screen,
                            normalize_frame)

OUT_OF_BATTLE_FRAMES = 30
# HUD 连续消失超过这么久，才认为真的离开了战斗、下次进关要从 0 起锚。
# 低于这个值说明只是部署视角/UI 过渡把 HUD 挡了一下，帧号必须保留。
RESET_AFTER_OUT_OF_BATTLE_S = 3.0


class WAACounter:
    def __init__(self, rate_percent: float = 100.0,
                 neg_override: bool | None = None,
                 neg_speed: int | None = None,
                 neg_declared: bool = True,
                 rate_candidates: list[float] | None = None,
                 require_hud: bool = False):
        self.rate_percent = float(rate_percent)
        self.neg_override = neg_override
        self.counter = WAAFrameCounter(self.rate_percent,
                                       neg_override=neg_override,
                                       neg_speed=neg_speed,
                                       neg_declared=neg_declared,
                                       rate_candidates=rate_candidates)
        # 回费速度操作：事件为 (帧号, 手速百分比或None(默认), 对应动作)。
        # 由计数线程在自己的 update 里按当前帧应用，主循环不用抢时序。
        self._rate_lock = threading.Lock()
        self._rate_events: list[tuple[int, float | None, object | None]] = []
        self.rate_callback = None
        self.base_rate_percent = float(rate_percent)
        self.base_normal_speed = self.counter.normal_speed
        self.base_required = self.counter.required
        self.base_eps = self.counter.eps
        self.base_neg_speed = self.counter.neg_speed
        self.base_neg_declared = self.counter.neg_declared
        self.armed = False
        self._arm_pending = False
        self._arm_frame = 0
        self._running: bool | None = None
        self.readable_streak = 0
        # 费用条和 HUD 都消失一段时间后视为离开战斗；下一次费用条出现时归零。
        self._out_of_battle_frames = 0
        self._out_of_battle = False
        self._out_of_battle_since = 0.0
        # 当前画面是否在关卡内；OCR 和正负费检测只在此时运行。
        self.in_battle = False
        # 正式运行时用右上角 HUD 按钮判断“在关卡内”，避免编队界面
        # 背景被误当成费用条；离线合成测试可以关掉。
        self.require_hud = bool(require_hud)
        self._t0 = time.perf_counter()
        self._n = 0
        self._csv_path: str | None = None
        self._csv_file = None
        self._csv_writer = None
        self._closed = False
        self._err: str | None = None

    # ---- 执行器用的接口 ----
    @property
    def frame(self) -> int:
        return self.counter.frame

    @property
    def frame_hi(self) -> int:
        return self.counter.frame + self.counter.uncertainty

    @property
    def last(self) -> BarVisual | None:
        return self.counter.last

    @property
    def csv_path(self) -> str | None:
        return self._csv_path

    @csv_path.setter
    def csv_path(self, path: str | None) -> None:
        self._csv_path = path

    def close_log(self) -> None:
        self._closed = True
        self.counter.close_log()
        if self._csv_file is not None:
            self._csv_file.close()
            self._csv_file = None
            self._csv_writer = None

    def arm(self, frame: int = 0) -> None:
        """等费用条画出来，把第一帧记为 frame。"""
        self.counter.reset(int(frame))
        self._arm_frame = int(frame)
        self._arm_pending = True
        self.armed = False
        self._out_of_battle_frames = 0
        self._out_of_battle = False
        self.in_battle = False

    def out_of_battle_s(self) -> float:
        """确认“离开战斗”后已经持续的秒数；仍在战斗（或没确认过）返回 0.0。

        给控制台的“退出关卡自动 stop”用：和 RESET_AFTER_OUT_OF_BATTLE_S
        同一个信号源——这个值一旦超过阈值，下次进关帧号也会归零。
        部署视角/UI 过渡造成的 HUD 短暂丢失（一秒左右）不会让它超过阈值。
        """
        if not self._out_of_battle:
            return 0.0
        since = getattr(self, "_out_of_battle_since", 0.0)
        return max(0.0, time.time() - since)

    def set_initial_frame(self, frame: int) -> None:
        self.arm(frame)

    def resync(self) -> None:
        self.arm(self.counter.frame)

    def advance_blind(self, frames: int = 1) -> int:
        """费用条静止时，由执行器确认“已盲走一帧”后推进模型。"""
        with self._rate_lock:
            return self.counter.advance_blind(frames)

    # ---- 回费速度操作 ----
    def set_rate_events(self, events) -> None:
        """设置按帧生效的回费速度事件，并先回到默认速度。"""
        with self._rate_lock:
            self._rate_events = list(events)
            self._set_rate_override(None)

    def rate_events_done(self) -> bool:
        with self._rate_lock:
            return not self._rate_events

    def _set_rate_override(self, value: float | None) -> None:
        """value=None 恢复默认/自动；数值为手动回费速度。

        调用时必须持有 _rate_lock（计数线程在 update 里调用）。
        """
        c = self.counter
        if value is None:
            c.rate_percent = self.base_rate_percent
            c.required = self.base_required
            c.eps = self.base_eps
            c.normal_speed = self.base_normal_speed
            c.neg_speed = self.base_neg_speed
            c.neg_declared = self.base_neg_declared
            c.manual_rate = None
            c.auto_detect = bool(c.rate_candidates)
            c._reset_pending()
            c._rate_samples.clear()
            c._cycle_pending_rate = None
            c._cycle_pending_hits = 0
        else:
            rate = float(value)
            normal = max(1, int(round(SPEED * rate / 100.0)))
            c.rate_percent = rate
            # required 固定为 2^24；回费速度只改每帧 speed。
            c.required = SCALE
            c.eps = EPS
            c.normal_speed = normal
            # 手动速度下，负费就是“当前速度 / 2”，不跟随设置里的第二档。
            c.neg_speed = max(1, normal // 2)
            c.neg_declared = True
            c.manual_rate = rate
            c.auto_detect = False
            c._reset_pending()
            c._rate_samples.clear()
            c._cycle_pending_rate = None
            c._cycle_pending_hits = 0
        c.speed = c.neg_speed if c.neg else c.normal_speed
        c.uncertainty = 0
        c.warn = ""
        # required 固定，acc 的相位表示不变；已经起锚时直接沿用。
        if self.armed and c.last is not None:
            c._need_anchor = False

    def _apply_rate_events_locked(self) -> None:
        """把已经到帧的事件应用到计数器；调用时需持有 _rate_lock。"""
        due = []
        while self._rate_events and self.counter.frame >= self._rate_events[0][0]:
            due.append(self._rate_events.pop(0))
        for _frame, value, action in due:
            self._set_rate_override(value)
            if action is not None and self.rate_callback is not None:
                try:
                    self.rate_callback(action)
                except Exception as e:
                    print(f"[rate] 回费速度回调失败：{e!r}")

    # ---- 逐帧 ----
    def update(self, frame_bgr) -> BarVisual | None:
        gray = normalize_frame(cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2GRAY))
        self._n += 1

        st = None
        try:
            st = read_buttons(gray)
        except Exception as e:
            if self._err is None:
                self._err = repr(e)
                print("[waa_counter] 按钮异常:", self._err)

        # 编队界面右下角的“开始行动”必须先排除：那里右上角还有
        # “快捷编队”按钮，白色文字会被暂停图标识别误判；背景也会被
        # 费用条检测误判。模板很快，足够每帧跑。
        start_screen = looks_like_start_screen(gray)
        # 费用条或 HUD 按钮还看得见，就仍算在战斗界面里。
        bar_visible = looks_like_cost_bar(gray) and not start_screen
        # 编队界面也可能有一个 2X 图标，不能只看 speed_drawn；
        # 真正战斗中的暂停/运行图标或子弹时间才是可靠信号。
        hud_visible = bool(not start_screen and st is not None and (
            st.pause_drawn or st.bullet or st.running is not None))
        body = gray[BAR_Y0:BAR_Y1, :]
        bar_has_detail = (body.size > 0
                          and float(body.max()) - float(body.min()) >= 5.0)
        # 纯色加载画面可能被 looks_like_cost_bar 误判；只用于“离开战斗”判断，
        # 不影响正常的费用条起锚。
        stage_bar_visible = bar_visible and bar_has_detail

        v: BarVisual | None = None
        # 计数线程独占计数器；主线程只通过 set_rate_events 换事件表。
        with self._rate_lock:
            if self.require_hud:
                # 正式运行：HUD 按钮是“在关卡内”的主信号。编队/加载界面
                # 也可能被费用条检测误判，不能只靠像素宽度起锚。
                if hud_visible:
                    if self._out_of_battle:
                        # 只有“离开战斗够久”才当成新关卡从 0 开始。
                        # 部署视角/UI 过渡时 HUD 会短暂消失一秒左右，
                        # 这时从 0 重新起锚会把帧号整个清掉
                        # （2026-11-09 1111 帧就是这么归零的）。
                        gap = time.time() - getattr(
                            self, "_out_of_battle_since", 0.0)
                        if gap >= RESET_AFTER_OUT_OF_BATTLE_S:
                            self.arm(0)
                        else:
                            self.resync()
                            print(f"[waa_counter] HUD 短暂丢失 {gap:.1f}s，"
                                  f"保留帧号 {self.counter.frame}")
                    self._out_of_battle_frames = 0
                    self._out_of_battle = False
                    in_battle_like = True
                else:
                    self._out_of_battle_frames += 1
                    # 战斗中 HUD 短暂丢失时，费用条仍可见就再给 30 帧宽限。
                    in_battle_like = bool(
                        self.armed and stage_bar_visible
                        and self._out_of_battle_frames < OUT_OF_BATTLE_FRAMES)
                    if self._out_of_battle_frames >= OUT_OF_BATTLE_FRAMES:
                        if not self._out_of_battle:
                            self._out_of_battle_since = time.time()
                        self._out_of_battle = True
            else:
                in_battle_like = stage_bar_visible or hud_visible
                if not in_battle_like:
                    self._out_of_battle_frames += 1
                    if self._out_of_battle_frames >= OUT_OF_BATTLE_FRAMES:
                        self._out_of_battle = True
                else:
                    # 之前确认离开战斗，现在费用条重新出现：新关卡从 0 开始。
                    if self._out_of_battle and stage_bar_visible:
                        self.arm(0)
                    self._out_of_battle_frames = 0
                    self._out_of_battle = False

            if not in_battle_like:
                # 不在关卡内：冻结，不跑费用条/正负费检测，也不做 OCR。
                v = self.counter.last
            elif self._arm_pending:
                if bar_visible:
                    self.counter.update(gray)      # 第一帧只建立基准
                    self.armed = True
                    self._arm_pending = False
                v = self.counter.last
            else:
                v = self.counter.update(gray)
            if in_battle_like:
                self._apply_rate_events_locked()
            self.in_battle = in_battle_like

        if st is not None:
            try:
                if st.running is not None:
                    self._running = st.running
                    self.readable_streak += 1
                else:
                    self.readable_streak = 0
                if v is not None and not self._out_of_battle:
                    v.running = self._running
                    v.m = st.m
                    v.bullet = st.bullet
                    v.pause_drawn = bool(st.pause_drawn)
                    v.speed_drawn = bool(st.speed_drawn)
                    v.mode = "bar"
                if self.armed and not self._out_of_battle:
                    self._log(v, st)
            except Exception as e:
                if self._err is None:
                    self._err = repr(e)
                    print("[waa_counter] 按钮/日志异常:", self._err)
        return v

    def _log(self, v: BarVisual | None, st) -> None:
        if self._closed or not self._csv_path:
            return
        if self._csv_file is None:
            self._csv_file = open(self._csv_path, "w", encoding="utf-8",
                                  newline="", buffering=1)
            self._csv_writer = csv.writer(self._csv_file)
            self._csv_writer.writerow([
                "idx", "t", "frame", "unc", "delta", "phase", "width",
                "cap_diff", "kind", "running", "m", "bullet", "neg",
                "pause_drawn", "speed_drawn", "warn",
            ])
        self._csv_writer.writerow([
            self._n, f"{time.perf_counter() - self._t0:.4f}",
            self.counter.frame, getattr(v, "uncertainty", 0),
            self.counter.last_delta,
            getattr(v, "phase", None), getattr(v, "width", None),
            f"{getattr(v, 'cap_diff', 0.0):.1f}",
            getattr(v, "kind", None),
            int(st.running) if st.running is not None else -1,
            st.m, int(st.bullet), int(getattr(v, "neg", False)),
            int(st.pause_drawn), int(st.speed_drawn),
            self.counter.warn,
        ])


class FrameCounterThread:
    """后台按顺序消费抓屏帧，喂给 WAACounter。"""

    def __init__(self, stream, counter: WAACounter, poll_s: float = 0.005):
        self.stream = stream
        self.counter = counter
        self.poll_s = poll_s
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self.updates = 0
        self.failures = 0

    def start(self) -> None:
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self) -> None:
        while not self._stop.is_set():
            frame = self.stream.get_frame(timeout=0.02)
            if frame is None:
                continue
            try:
                self.counter.update(frame)
                self.updates += 1
            except Exception:
                self.failures += 1

    def stop(self) -> None:
        self._stop.set()
        # 必须等线程真正退出再关日志；否则线程手里那一帧会在 close 之后
        # 重新以 "w" 打开文件，把整场的 CSV 截成一行。
        if self._thread is not None:
            self._thread.join(timeout=0.5)
