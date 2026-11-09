"""离线测试：同帧不 resume、选中确认、E/Q 无反应重试。"""

from __future__ import annotations

import os
import sys
import threading

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from types import SimpleNamespace  # noqa: E402

from waa.core.executor import BULLET_STEP_PERIOD, Executor  # noqa: E402


class FakePause(Executor):
    def __init__(self):
        self.frame = 240
        self.paused = True
        self.resumed = False
        self.paused_calls = 0
        self.pause_samples = []
        self.pause_method = "touch"
        self.speed_override = None
        self.counter = SimpleNamespace(
            frame_hi=240,
            last=SimpleNamespace(mode="bar", warn="", m=1.0),
        )

    def current_frame(self):
        return self.frame

    def is_paused(self):
        return self.paused

    def resume(self, *a, **k):
        self.resumed = True
        self.paused = False

    def pause(self, *a, **k):
        self.paused_calls += 1
        self.paused = True

    def _settled_frame(self, *a, **k):
        return self.frame


class FakeKey(Executor):
    def __init__(self, states):
        self.states = list(states)
        self.state = True
        self.sent = 0
        self.steps = 0
        self.selects = 0
        self.frame = 100
        self.client_left = self.client_top = 0
        self.cw, self.ch = 1280, 720
        self.exit_clicks = 0

        class Touch:
            def __init__(self, owner):
                self.owner = owner

            def click(self, *a, **k):
                self.owner.exit_clicks += 1

        self.touch = Touch(self)

    def _send_action_key(self, vk, label):
        if self.sent >= len(self.states):
            self.state = True
        else:
            self.state = bool(self.states[self.sent])
        self.sent += 1
        print(f"  fake send {label} -> bullet={self.state}")

    def bullet_state(self):
        return self.state

    def _wait_bullet(self, want, timeout=0.6):
        return self.state == want

    def current_frame(self):
        return self.frame

    def _step_one_bullet(self, action, label):
        self.steps += 1
        self.frame += 1
        print(f"  fake step {label} -> frame={self.frame}")
        return self.frame

    def _select_operator(self, action, allow_late=False):
        self.selects += 1
        self.state = True
        print(f"  fake reselect allow_late={allow_late}")
        return "fake"


class FakeDrag(Executor):
    def __init__(self, event):
        self.manual_pause = event
        self.downs = []
        self.moves = []
        self.ups = []

        class Touch:
            def __init__(self, owner):
                self.owner = owner

            def touch_down(self, *p):
                self.owner.downs.append(p)

            def touch_move(self, *p):
                self.owner.moves.append(p)

            def touch_up(self, *p):
                self.owner.ups.append(p)

        self.touch = Touch(self)


class FakeStep(Executor):
    def __init__(self, speed):
        self.speed = speed
        self.period = None

    def _speed(self):
        return self.speed

    def step_one_frame_afa(self, period, require_bullet=False,
                           allow_blind=False, allow_overshoot=False):
        self.period = period
        return 1


class FakeDeployOvershoot(Executor):
    def __init__(self):
        self.step_enabled = True
        self.frame = 31
        self.cancel_clicks = 0
        self.pause_method = "touch"
        self.speed_override = 2

        class Touch:
            def __init__(self, owner):
                self.owner = owner

            def click(self, *a, **k):
                self.owner.cancel_clicks += 1

        self.touch = Touch(self)
        self.stream = SimpleNamespace(latest=lambda: None)
        self.counter = SimpleNamespace(last=SimpleNamespace(bullet=True))

    def _click_rightmost_deploy_card(self):
        pass

    def bullet_state(self):
        return True

    def current_frame(self):
        return self.frame

    def screen(self, x, y):
        return 0, 0

    def _speed(self):
        return 2.0


class FakeSelect(Executor):
    def __init__(self, results, frame=600):
        self.results = results
        self.calls = []
        self.frame = frame
        self.paused = True
        self.bullet = False
        self.select_pause_mode = "tap"

    def current_frame(self):
        return self.frame

    def is_paused(self):
        return self.paused

    def pause(self, *a, **k):
        self.paused = True

    def bullet_state(self):
        return self.bullet

    def _wait_bullet(self, want, timeout=0.6):
        return self.bullet == want

    def select_while_paused(self, rx, ry, pause_mode="tap"):
        self.calls.append(pause_mode)
        self.bullet = bool(self.results.get(pause_mode, False))
        print(f"  fake select mode={pause_mode} -> bullet={self.bullet}")

    def _enter_bullet(self):
        self.calls.append("fallback")
        self.bullet = True
        print("  fake enter bullet fallback")
        return True


def check(cond, label):
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def main() -> int:
    checks = []

    # 1) 已经在 240 暂停，目标 238：不能再 resume。
    ex = FakePause()
    ret = Executor.pause_at(ex, 238)
    checks.append(check(ret == 240, f"already paused returns {ret}"))
    checks.append(check(not ex.resumed, "already paused does not resume"))

    action = SimpleNamespace(view_front=(0.5, 0.5), frame=100)

    # 2) 成功：按一次后 bullet=False。
    ex2 = FakeKey([False])
    ok = Executor._send_action_verified(
        ex2, action, 0x51, "撤退 Q", "撤退")
    checks.append(check(ok is True, f"success returned {ok}"))
    checks.append(check(ex2.sent == 1 and ex2.steps == 0,
                        f"success sent={ex2.sent} steps={ex2.steps}"))

    # 3) 技能/撤退释放失败：首次 + 最多 3 次重试，仍失败就点退出点。
    for vk, label, kind in ((0x45, "技能 E", "技能"),
                            (0x51, "撤退 Q", "撤退")):
        exf = FakeKey([True, True, True, True])
        okf = Executor._send_action_verified(exf, action, vk, label, kind)
        checks.append(check(okf is False, f"{kind} failure returned {okf}"))
        checks.append(check(exf.sent == 3 and exf.steps == 0,
                            f"{kind} failure sent={exf.sent} "
                            f"steps={exf.steps}"))
        checks.append(check(exf.exit_clicks == 1,
                            f"{kind} exit click = {exf.exit_clicks}"))

    # 4) 601 这种相邻帧：先做一次 direct 前置轻点（唤醒客户端），再走三连点。
    ex4 = FakeSelect({"direct": True})
    action = SimpleNamespace(view_front=(0.5, 0.5), frame=601)
    mode = Executor._select_operator(ex4, action)
    checks.append(check(mode == "direct", f"selection mode = {mode}"))
    checks.append(check(ex4.calls[:1] == ["direct"],
                        f"direct tried first: {ex4.calls}"))

    # 5) direct 不中时回退 tap，并且 tap 成功。
    ex5 = FakeSelect({"direct": False, "tap": True})
    mode5 = Executor._select_operator(ex5, action)
    checks.append(check(mode5 == "tap", f"fallback mode = {mode5}"))
    checks.append(check(ex5.calls[:2] == ["direct", "tap"],
                        f"fallback order: {ex5.calls}"))

    # 6) 一直选不中要报错。
    ex6 = FakeSelect({})
    try:
        Executor._select_operator(ex6, action)
        checks.append(check(False, "selection fail should raise"))
    except RuntimeError as e:
        checks.append(check(True, f"selection fail raised: {e}"))

    # 7) 子弹时间一帧不随倍速变化：1X / 2X 都用 5 × 69905 / 2^21。
    for speed in (1.0, 2.0):
        ex7 = FakeStep(speed)
        Executor.step_one_frame(ex7, require_bullet=True)
        checks.append(check(abs(ex7.period - BULLET_STEP_PERIOD) < 1e-12,
                            f"{speed:g}X step period = {ex7.period}"))

    # 8) 用户要求 2X lead 保持 5。
    ex_lead = FakeDeployOvershoot()
    checks.append(check(Executor._default_lead(ex_lead) == 5,
                        f"2X lead = {Executor._default_lead(ex_lead)}"))

    # 9) 部署已经过帧时不再拖拽，直接判失败并取消选中。
    overshoot_action = SimpleNamespace(
        oper="斑点", view_side=(0.5, 0.5), direction="右", frame=30)
    ret = Executor.deploy(ex_lead, overshoot_action)
    checks.append(check(ret is None,
                        f"deploy overshoot returns {ret}"))
    checks.append(check(ex_lead.cancel_clicks == 1,
                        f"deploy overshoot cancel clicks = "
                        f"{ex_lead.cancel_clicks}"))

    # 10) 拖拽中按 F6：立即松手，不能松在目标格。
    ev = threading.Event()
    ex_drag = FakeDrag(ev)
    t = threading.Timer(0.02, ev.set)
    t.start()
    interrupted = Executor._drag_card_to_tile(
        ex_drag, (100, 600), (150, 500), (300, 400))
    t.cancel()
    checks.append(check(interrupted is True,
                        f"drag F6 interrupted = {interrupted}"))
    checks.append(check(ex_drag.downs == [(100, 600)],
                        f"drag down = {ex_drag.downs}"))
    checks.append(check(ex_drag.ups == [(100, 600)],
                        f"drag up = {ex_drag.ups}"))
    ev.clear()

    # 11) 方向拖拽中按 F6：不立即松手，按原速拖回起点再松手。
    ev2 = threading.Event()
    ex_dir = FakeDrag(ev2)
    t2 = threading.Timer(0.22, ev2.set)
    t2.start()
    interrupted2 = Executor._drag_direction(
        ex_dir, (100, 600), 50, 0)
    t2.cancel()
    checks.append(check(interrupted2 is True,
                        f"direction F6 interrupted = {interrupted2}"))
    checks.append(check(ex_dir.downs == [(100, 600)],
                        f"direction down = {ex_dir.downs}"))
    checks.append(check(ex_dir.ups == [(100, 600)],
                        f"direction up at start = {ex_dir.ups}"))
    checks.append(check(len(ex_dir.moves) >= 2,
                        f"direction move count = {len(ex_dir.moves)}"))
    ev2.clear()

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
