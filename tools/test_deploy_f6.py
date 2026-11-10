"""离线测试：部署落点后按 F6 → 取消子弹 → 恢复后仍在指定帧 → 重新部署。"""

from __future__ import annotations

import os
import sys
import threading
from types import SimpleNamespace

import numpy as np

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import waa.core.executor as executor_mod  # noqa: E402
from waa.core.executor import Executor  # noqa: E402


class FakeDeployF6(Executor):
    def __init__(self):
        self.manual_pause = threading.Event()
        self.deploy_active = False
        self.frame = 30
        self.step_enabled = True
        self.cancels = 0
        self.holds = 0
        self.drag_calls = 0
        self.direction_calls = 0
        self.cw, self.ch = 1280, 720
        self.client_left = self.client_top = 0
        self.counter = SimpleNamespace(
            last=SimpleNamespace(t=0.0, mode="bar", warn="", bullet=True))
        self.stream = SimpleNamespace(
            latest=lambda: np.zeros((720, 1280, 3), dtype=np.uint8))
        self.touch = SimpleNamespace(click=lambda *a, **k: None)

    def _click_rightmost_deploy_card(self):
        pass

    def current_frame(self):
        return self.frame

    def bullet_state(self):
        return True

    def screen(self, x, y):
        return int(x), int(y)

    def _drag_card_to_tile(self, card, middle, drop):
        self.drag_calls += 1
        if self.drag_calls == 1:
            self.manual_pause.set()
        return False

    def _drag_direction(self, tile, dx, dy):
        self.direction_calls += 1
        return False

    def _cancel_bullet_time(self):
        self.cancels += 1
        print("  fake cancel bullet")

    def _manual_hold(self):
        self.holds += 1
        print("  fake manual hold clears F6")
        self.manual_pause.clear()
        return 0.0


def check(cond: bool, label: str) -> bool:
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def main() -> int:
    original = executor_mod.locate_avatar
    executor_mod.locate_avatar = lambda frame, oper: (0.5, 0.9, 1.0)
    try:
        ex = FakeDeployF6()
        action = SimpleNamespace(
            oper="斑点", view_side=(0.5, 0.5), direction="右", frame=30)
        ret = Executor.deploy(ex, action)
    finally:
        executor_mod.locate_avatar = original

    checks = [
        check(ret == 30, f"redo deploy returns {ret}"),
        check(ex.cancels == 1, f"cancel bullet clicks = {ex.cancels}"),
        check(ex.holds == 1, f"manual holds = {ex.holds}"),
        check(ex.drag_calls == 2, f"drag calls = {ex.drag_calls}"),
    ]
    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
