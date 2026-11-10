"""离线测试：技能/撤退新流程（右卡进子弹 → 步进 → 点场上目标 → E/Q）。"""

from __future__ import annotations

import os
import sys
import threading
from types import SimpleNamespace

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa.core.executor import Executor  # noqa: E402


class FakeReporter:
    def __init__(self):
        self.dones = []

    def action_start(self, idx, action):
        pass

    def action_done(self, idx, action, target, actual, note):
        self.dones.append((idx, target, actual, note))


class FakeBatch(Executor):
    def __init__(self, frame):
        self.manual_pause = threading.Event()
        self.reporter = FakeReporter()
        self.current_action = None
        self.step_enabled = True
        self.frame = frame
        self.events = []
        self.counter = SimpleNamespace(
            last=SimpleNamespace(t=0.0, mode="bar", warn=""))
        self._last_deploy_card_screen = (10, 20)

        def click(*a, **k):
            self.events.append(("click", a[0], a[1]))

        self.touch = SimpleNamespace(click=click)

    def current_frame(self):
        return self.frame

    def _enter_bullet_by_card(self):
        self.events.append("card")
        return True

    def _select_operator(self, action):
        self.events.append("fallback_select")
        return "fallback"

    def step_one_frame(self, *a, **k):
        self.events.append("step")
        self.frame += 1
        return 1

    def select_while_paused(self, rx, ry, pause_mode="hold"):
        self.events.append(("target", pause_mode))
        return None

    def bullet_state(self):
        return True

    def _send_action_verified(self, action, vk, label, kind):
        self.events.append("eq")
        return True

    def is_paused(self):
        return True

    def screen(self, x, y):
        return 0, 0


def check(cond: bool, label: str) -> bool:
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def main() -> int:
    checks = []
    action = SimpleNamespace(
        kind="技能", oper="斑点", view_front=(0.5, 0.5), frame=100,
        alias=None, to_string=None)
    ex = FakeBatch(frame=100)
    Executor.batch_skill_retreat(ex, [action], start_index=0,
                                 resume_after=False)
    checks.append(check(ex.events == [
        "card", ("click", 10, 20), ("target", "direct"), "eq",
        ("click", 0, 0)],
        f"事件顺序 = {ex.events}"))
    checks.append(check(ex.reporter.dones
                        and ex.reporter.dones[0][1:] == (100, 100, "OK"),
                        f"结果 = {ex.reporter.dones}"))

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
