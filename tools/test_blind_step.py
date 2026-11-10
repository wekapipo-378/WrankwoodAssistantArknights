"""离线测试：负费固定 /2、低回费盲步进。"""

from __future__ import annotations

import os
import sys
import threading
from types import SimpleNamespace

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa.core.executor import Executor  # noqa: E402
from waa.core.frame_counter import WAAFrameCounter  # noqa: E402
from tools.run_axis import _resolve_rate_config  # noqa: E402


def check(cond, label):
    print(("PASS " if cond else "FAIL ") + label)
    return bool(cond)


class FakeFrameSource:
    """费用条静止：frame 不动，但允许盲步进推进模型。"""

    def __init__(self, frame=33, uncertainty=7):
        self.frame = frame
        self.uncertainty = uncertainty
        self.last = SimpleNamespace(bullet=True)
        self.advanced = 0

    @property
    def frame_hi(self):
        return self.frame + self.uncertainty

    def advance_blind(self, frames=1):
        self.advanced += int(frames)
        self.frame += int(frames)
        self.uncertainty = max(0, self.uncertainty - int(frames))
        return self.frame


class FakePrepare(Executor):
    def __init__(self, source, blind=True):
        self.counter = source
        self.axis = SimpleNamespace(
            settings={"低回费盲步进": blind,
                      # 安全下限去掉后，lead 完全由轴设置决定；夹具要显式给出
                      # 提前量，否则默认 5 帧会让 41-33=8 > 5 走进 pause_at，
                      # 而这个替身没有 stream/resume。190ms 在 2X 下 = 11 帧，
                      # 与旧下限的取值一致，保持这组用例原来的分支。
                      "1x操作提前量": "190ms",
                      "2x操作提前量": "190ms"})
        self.blind_flags = []
        self.last_pause_trigger = 0
        self.last_pause_landing = 0
        # 暂停样本不再参与 lead 计算，保留空列表只为兼容旧接口。
        self.pause_samples = []
        self.speed_override = 2

    def current_frame(self):
        return self.counter.frame

    def current_frame_hi(self):
        return self.counter.frame_hi

    def _default_lead(self):
        return 10

    def is_paused(self):
        return True

    def is_running_icon(self):
        return False

    def bullet_state(self):
        return True

    def _log_pause(self, *a, **k):
        pass

    def step_one_frame(self, require_bullet=False, allow_blind=False, **k):
        self.blind_flags.append(bool(allow_blind))
        if not allow_blind:
            raise RuntimeError("test: 未允许盲步进")
        self.counter.advance_blind(1)


class FakeStep(Executor):
    def __init__(self, source):
        self.counter = source
        self.manual_pause = threading.Event()
        self.client_left = self.client_top = 0
        self.cw, self.ch = 1280, 720
        self._last_icon_score = 0.99
        self.taps = 0
        self.clicks = 0
        self._paused = True

        class Touch:
            def __init__(self, owner):
                self.owner = owner

            def tap(self, *a, **k):
                self.owner.taps += 1

            def click(self, *a, **k):
                self.owner.clicks += 1

        self.touch = Touch(self)

    def current_frame(self):
        return self.counter.frame

    def current_frame_hi(self):
        return self.counter.frame_hi

    def is_paused(self):
        return self._paused

    def bullet_state(self):
        return True

    def _settled_frame(self, *a, **k):
        # 条没动：计数器仍停在原帧。
        return self.current_frame()

    def _sleep_precise(self, seconds):
        pass

    def screen(self, rx, ry):
        return (int(rx * self.cw), int(ry * self.ch))


def main() -> int:
    checks = []

    # 1) 负费固定 = 当前正常速度 // 2，不读第二档。
    axis = SimpleNamespace(settings={"回费速度": [1, 1.1, 0.5]})
    base, neg, declared, src, rates = _resolve_rate_config(axis, None)
    checks.append(check(base == 100.0, f"基准 = {base:g}%"))
    checks.append(check(neg == 279620, f"100% 负费 speed = {neg}"))
    checks.append(check(declared is True, "负费档始终视为已知"))
    checks.append(check(rates == [100.0, 110.0, 50.0], f"读到三档 {rates}"))

    # 2) 模型盲步进：帧号和 acc 都按当前 speed 前进。
    c = WAAFrameCounter(100)
    c.reset(0)
    c.advance_blind(3)
    checks.append(check(c.frame == 3, f"盲走 3 帧后 frame={c.frame}"))
    checks.append(check(c.acc == 3 * c.normal_speed,
                        f"acc 前进三帧 = {c.acc}"))

    # 3) 目标落在同宽区间内：_prepare_to_frame 连走 3 次盲步进。
    prep_afc = FakeFrameSource(frame=33, uncertainty=7)
    prep = FakePrepare(prep_afc, blind=True)
    ok = prep._prepare_to_frame(None, 36)
    checks.append(check(ok and prep_afc.frame == 36,
                        f"33~40 同宽，目标 36 -> frame={prep_afc.frame}"))
    checks.append(check(prep.blind_flags == [True, True, True],
                        f"三次都允许盲走 {prep.blind_flags}"))

    # 4) 轴设置关闭盲步进：不允许盲走。
    prep2 = FakePrepare(FakeFrameSource(frame=33, uncertainty=7), blind=False)
    raised = False
    try:
        prep2._prepare_to_frame(None, 36)
    except RuntimeError:
        raised = True
    checks.append(check(raised and prep2.blind_flags == [False],
                        f"设置 FALSE 时不盲走 {prep2.blind_flags}"))

    # 5) 目标超过上界：不盲走（没有依据）。
    prep3 = FakePrepare(FakeFrameSource(frame=33, uncertainty=7), blind=True)
    raised = False
    try:
        prep3._prepare_to_frame(None, 41)
    except RuntimeError:
        raised = True
    checks.append(check(raised and prep3.blind_flags == [False],
                        f"目标 41 > hi 40 时不盲走 {prep3.blind_flags}"))

    # 6) 可盲走时：0 帧变化也算 1 帧，不报错。
    source = FakeFrameSource()
    ex = FakeStep(source)
    got = ex.step_one_frame_afa(period=0, require_bullet=True,
                                allow_blind=True)
    checks.append(check(got == 1 and source.frame == 34,
                        f"盲步进 33->{source.frame}"))
    checks.append(check(source.advanced == 1, "只盲走一次"))
    checks.append(check(ex.taps + ex.clicks == 2,
                        "一次步进只发左继续+右暂停两下"))

    # 4) 不允许盲走时：连续 0 帧仍按原规则重试 5 次后报错。
    afc2 = FakeFrameSource()
    ex2 = FakeStep(afc2)
    raised = False
    try:
        ex2.step_one_frame_afa(period=0, require_bullet=True,
                               allow_blind=False)
    except RuntimeError:
        raised = True
    checks.append(check(raised, "关闭盲步进后 0 帧会报错"))
    checks.append(check(ex2.taps + ex2.clicks == 10,
                        f"5 次重试共 {ex2.taps + ex2.clicks} 次点击"))

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
