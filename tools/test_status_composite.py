"""离线测试：地图页“当前状态”的倍速/暂停/子弹/选中单位组合。"""

from __future__ import annotations

import os
import sys
import threading
from types import SimpleNamespace

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa.core.excel_report import ExcelReporter  # noqa: E402


class FakeRuntime:
    def __init__(self, m=1.0, running=True, bullet=False):
        self.frame = 123
        self.frame_hi = 123
        self.counter = SimpleNamespace(neg=False, spec_error="")
        self.last = SimpleNamespace(m=m, running=running, bullet=bullet)


class FakeEx:
    def __init__(self):
        self.manual_pause = threading.Event()
        self.manual_resume_wait = False
        self.current_action = None
        self.selected_oper = None
        self.phase = "运行中"


def make_reporter(runtime, ex):
    axis = SimpleNamespace(path=os.path.join(ROOT, "axes", "WAA_标准模板.xlsx"))
    rep = ExcelReporter(axis, {"base": 100.0, "declared": [100.0]},
                        os.path.join(ROOT, "out", "_status_test"),
                        enabled=False)
    rep._counter = runtime
    rep._executor = ex
    rep._poll_runtime()
    return rep


def check(cond: bool, label: str) -> bool:
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def main() -> int:
    checks = []
    ex = FakeEx()
    ex.selected_oper = "斑点"
    rep = make_reporter(FakeRuntime(m=2.0, running=False, bullet=True), ex)
    checks.append(check(rep._speed_text == "0.2X"
                        and rep._selected_text == "斑点"
                        and rep._status == "运行中",
                        f"子弹时间 = {rep._speed_text}/{rep._selected_text}/{rep._status}"))

    ex2 = FakeEx()
    rep2 = make_reporter(FakeRuntime(m=1.0, running=True, bullet=False), ex2)
    checks.append(check(rep2._speed_text == "1X"
                        and rep2._selected_text == ""
                        and rep2._status == "运行中",
                        f"1X/运行 = {rep2._speed_text}/{rep2._status}"))

    ex3 = FakeEx()
    ex3.phase = "手动暂停中"
    rep3 = make_reporter(FakeRuntime(m=2.0, running=False, bullet=False), ex3)
    checks.append(check(rep3._speed_text == "2X"
                        and rep3._status == "手动暂停中",
                        f"手动暂停 = {rep3._speed_text}/{rep3._status}"))

    ex4 = FakeEx()
    ex4.phase = "暂停操作：等待手动恢复"
    rep4 = make_reporter(FakeRuntime(m=1.0, running=False, bullet=False), ex4)
    checks.append(check(rep4._speed_text == "1X"
                        and rep4._status == "暂停操作：等待手动恢复",
                        f"暂停操作等待 = {rep4._speed_text}/{rep4._status}"))

    ex5 = FakeEx()
    rep5 = make_reporter(FakeRuntime(m=1.0, running=True, bullet=False), ex5)
    rep5._finished = True
    rep5._ok = True
    rep5._poll_runtime()
    checks.append(check(rep5._status == "轴完成",
                        f"轴完成后状态 = {rep5._status}"))

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
