"""离线测试：轴里的“暂停”操作。"""

from __future__ import annotations

import os
import shutil
import sys
import threading
import time
from types import SimpleNamespace

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from openpyxl import load_workbook  # noqa: E402

from waa.core.axis import AxisAction, read_axis  # noqa: E402
from waa.core.executor import Executor  # noqa: E402

TMP = os.path.join(ROOT, "out", "_pause_test")
SRC = os.path.join(ROOT, "axes", "WAA_标准模板.xlsx")
DST = os.path.join(TMP, "WAA_pause_test.xlsx")


class FakeHold(Executor):
    def __init__(self):
        self.manual_pause = threading.Event()
        self.manual_resume_wait = False
        self.running = False
        self.resumed = 0
        self.reporter = FakeReporter()
        self.axis_running = True  # wait_manual_resume 是轴运行中的场景

    def is_running_icon(self):
        return self.running

    def resume(self, *a, **k):
        self.resumed += 1
        self.running = True

    def _focus_game(self):
        return True

    def reload_pending_actions(self):
        return False

    def _park_cursor(self):
        pass


class FakeReporter:
    def __init__(self):
        self.starts = []
        self.dones = []
        self.enabled = True
        self.started = True
        self.fronts = 0

    def action_start(self, idx, action):
        self.starts.append(idx)

    def bring_to_front(self):
        self.fronts += 1
        return True

    def action_done(self, idx, action, target, actual, note):
        self.dones.append((idx, target, actual, note))


class FakeRun(Executor):
    def __init__(self, action):
        self.axis = SimpleNamespace(actions=[action])
        self.manual_pause = threading.Event()
        self.manual_resume_wait = False
        self.reporter = FakeReporter()
        self.current_action = None
        self.pending_from = 0
        self.end_pause = False
        self.waited = False
        self.resumed = 0
        self.frame = 98
        self.bullet = True
        self.pause_targets = []
        self.stepped = 0
        self.exited_card = 0
        self.selected_oper = None
        self._last_deploy_card_screen = (10, 20)
        # run() 起手会按已部署栈重建状态；替身给空栈。
        self.deployed_stack = []
        self.stop_requested = threading.Event()

        def click(x, y, **kwargs):
            if (x, y) == self._last_deploy_card_screen:
                self.exited_card += 1
                self.bullet = False

        self.touch = SimpleNamespace(click=click)

    def pause_at(self, target):
        print(f"  fake pause_at {target}")
        self.pause_targets.append(target)
        return self.frame

    def _enter_bullet_by_card(self):
        self.bullet = True
        return True

    def step_one_frame(self, *a, **k):
        self.stepped += 1
        self.frame += 1
        return 1

    def current_frame(self):
        return self.frame

    def bullet_state(self):
        return self.bullet

    def _sleep_or_f6(self, seconds):
        return False

    def is_paused(self):
        return True

    def is_running_icon(self):
        return False

    def wait_manual_resume(self):
        self.waited = True

    def _prepare_to_frame(self, action, target):
        self.frame = target
        self.bullet = True
        return True

    def _exit_bullet_chain(self):
        self.bullet = False
        return True

    def resume(self, *a, **k):
        self.resumed += 1

    def _cancel_bullet_time(self):
        self.bullet = False


def check(cond: bool, label: str) -> bool:
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def main() -> int:
    os.makedirs(TMP, exist_ok=True)
    if os.path.exists(DST):
        os.remove(DST)
    shutil.copy2(SRC, DST)
    wb = load_workbook(DST)
    ws = wb["作战记录"]
    ws["B2"] = 100
    ws["D2"] = "暂停"
    wb.save(DST)
    wb.close()

    checks = []
    axis = read_axis(DST)
    checks.append(check(len(axis.actions) == 1,
                        f"暂停行被读入 actions={len(axis.actions)}"))
    checks.append(check(not axis.warnings, f"暂停行没有警告 {axis.warnings}"))
    if axis.actions:
        a = axis.actions[0]
        checks.append(check(a.frame == 100 and a.kind == "暂停"
                            and a.tile_pos is None,
                            f"暂停 action frame={a.frame} kind={a.kind} "
                            f"tile={a.tile_pos}"))

    # 手动点游戏暂停键恢复
    ex = FakeHold()
    t = threading.Thread(target=Executor.wait_manual_resume, args=(ex,))
    t.start()
    time.sleep(0.15)
    ex.running = True
    t.join(timeout=3)
    checks.append(check(not t.is_alive() and ex.resumed == 0,
                        f"点击游戏恢复：resumed={ex.resumed} alive={t.is_alive()}"))
    checks.append(check(ex.manual_resume_wait is False,
                        "等待标记已清除（点击恢复）"))

    # F6：第一次前置 Excel 保持暂停，第二次恢复
    ex2 = FakeHold()
    t2 = threading.Thread(target=Executor.wait_manual_resume, args=(ex2,))
    t2.start()
    time.sleep(0.15)
    ex2.manual_pause.set()
    time.sleep(0.2)
    checks.append(check(t2.is_alive() and ex2.resumed == 0,
                        "第一次 F6：保持暂停、不恢复"))
    checks.append(check(ex2.reporter.fronts == 1,
                        f"第一次 F6：Excel 已前置 {ex2.reporter.fronts} 次"))
    ex2.manual_pause.clear()
    t2.join(timeout=3)
    checks.append(check(not t2.is_alive() and ex2.resumed == 1,
                        f"第二次 F6 恢复：resumed={ex2.resumed} "
                        f"alive={t2.is_alive()}"))
    checks.append(check(not ex2.manual_pause.is_set(),
                        "F6 事件已清除"))

    # 主循环：跑到目标帧 → 退出子弹 → 等手动恢复
    run = FakeRun(AxisAction(frame=100, kind="暂停"))
    Executor.run(run, max_actions=1)
    checks.append(check(run.frame == 100, f"停在目标帧 {run.frame}"))
    checks.append(check(run.waited, "暂停后进入等待手动恢复"))
    checks.append(check(run.reporter.dones
                        and run.reporter.dones[0][1:] == ("100", 100, "成功"),
                        f"暂停结果 = {run.reporter.dones}"))

    # pause_by_key（开局暂停）：照搬 AFA——等倍速按钮区域的「纯白像素」出现就
    # 发一次 Space，不确认。用 speed_white_pixel 这个精确信号，不用相对亮度的
    # speed_drawn（后者在加载阶段误判）。
    import numpy as np
    import waa.core.buttons as bt

    class FakePauseByKey(Executor):
        def __init__(self):
            self.tap_times = []
            self.paused = False
            self._paused_at = 0.0
            self._white = False   # 模拟倍速按钮白像素是否出现
            self._frame = np.zeros((1080, 1920, 3), dtype=np.uint8)  # BGR 三通道
            # fake stream：返回一个全零 BGR 帧（mock 掉 speed_white_pixel 后不关心内容）
            self.stream = SimpleNamespace(latest=lambda: self._frame)

        def _tap_space(self):
            self.tap_times.append(time.perf_counter())

    # 场景1：白像素出现 → 发一次 Space → 返回 True（不确认）
    pk = FakePauseByKey()
    pk._white = True
    orig_speed_white = bt.speed_white_pixel
    bt.speed_white_pixel = lambda gray: pk._white
    try:
        ok = pk.pause_by_key(wait_armed=0.5)
    finally:
        bt.speed_white_pixel = orig_speed_white
    checks.append(check(ok and len(pk.tap_times) == 1 and pk.paused,
                        f"白像素出现发一次 Space 即返回 "
                        f"（按了 {len(pk.tap_times)} 次）"))

    # 场景2：白像素一直没出现（加载超时）→ 不按键、返回 False
    pk2 = FakePauseByKey()
    pk2._white = False
    orig_speed_white = bt.speed_white_pixel
    bt.speed_white_pixel = lambda gray: pk2._white
    try:
        ok2 = pk2.pause_by_key(wait_armed=0.3)
    finally:
        bt.speed_white_pixel = orig_speed_white
    checks.append(check(not ok2 and len(pk2.tap_times) == 0 and not pk2.paused,
                        f"没等到白像素不按键不置 paused "
                        f"（按了 {len(pk2.tap_times)} 次）"))

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
