"""离线测试：F6 在“动作暂停 / 手动暂停 / 暂停操作等待”下的行为。"""

from __future__ import annotations

import os
import sys
import threading

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa.core.executor import Executor  # noqa: E402


class FakeF6(Executor):
    def __init__(self, running: bool):
        self.manual_pause = threading.Event()
        self.manual_resume_wait = False
        self.deploy_active = False
        self.running = running
        self.resumed = 0
        self.paused = 0

    def is_running_icon(self):
        return self.running

    def resume(self, *a, **k):
        self.resumed += 1
        self.running = True

    def pause(self, *a, **k):
        self.paused += 1
        self.running = False


def check(cond: bool, label: str) -> bool:
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def main() -> int:
    checks = []

    # 动作暂停中按 F6：只登记请求，不能被热键线程直接恢复。
    ex = FakeF6(running=False)
    Executor.toggle_manual_pause(ex)
    checks.append(check(ex.manual_pause.is_set(),
                        "动作暂停：F6 登记暂停请求"))
    checks.append(check(ex.resumed == 0 and ex.paused == 0,
                        "动作暂停：不直接 resume/pause"))

    # 再按一次 F6：只清除请求，恢复交给主线程安全点。
    Executor.toggle_manual_pause(ex)
    checks.append(check(not ex.manual_pause.is_set(),
                        "第二次 F6：清除暂停请求"))
    checks.append(check(ex.resumed == 0,
                        "第二次 F6：不直接 resume"))

    # 暂停操作等待中按 F6：第一次只登记（主线程前置 Excel），不恢复。
    ex2 = FakeF6(running=False)
    ex2.manual_resume_wait = True
    Executor.toggle_manual_pause(ex2)
    checks.append(check(ex2.manual_pause.is_set() and ex2.resumed == 0,
                        "暂停操作等待：第一次 F6 登记请求、不恢复"))
    Executor.toggle_manual_pause(ex2)
    checks.append(check(not ex2.manual_pause.is_set() and ex2.resumed == 0,
                        "暂停操作等待：第二次 F6 清除请求，恢复交给主线程"))

    # 运行中按 F6：只登记请求；暂停由主线程 _manual_hold 单点完成，
    # 避免热键线程和主线程各点一次切换按钮互相抵消。
    ex3 = FakeF6(running=True)
    Executor.toggle_manual_pause(ex3)
    checks.append(check(ex3.manual_pause.is_set() and ex3.paused == 0,
                        "运行中：F6 只登记请求，不在热键线程点暂停"))

    # 部署拖拽中按 F6：只登记，立即暂停交给部署流程处理。
    ex4 = FakeF6(running=True)
    ex4.deploy_active = True
    Executor.toggle_manual_pause(ex4)
    checks.append(check(ex4.manual_pause.is_set() and ex4.paused == 0,
                        "部署中：F6 只登记，不抢先暂停"))

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
