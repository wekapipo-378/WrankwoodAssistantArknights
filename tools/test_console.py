"""离线测试：F5 独立 write 文件、offset/frame、MarkWriter、帧窗口透明度。"""

from __future__ import annotations

import cv2
import json
import numpy as np
import os
import shutil
import sys
import threading
import time
from types import SimpleNamespace

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from openpyxl import load_workbook  # noqa: E402

from synth_neg_test import build  # noqa: E402
from tools.axis_console import (  # noqa: E402
    ConsoleSession, ShowFrameWindow, _make_preview_counter, _parse_alpha,
)
from waa.core.axis import AxisAction  # noqa: E402
from waa.core.excel_report import ExcelReporter  # noqa: E402
from waa.core.executor import Executor  # noqa: E402
from waa.core.mark_writer import MarkWriter  # noqa: E402

TMP = os.path.join(ROOT, "out", "_console_test")
SRC = os.path.join(ROOT, "axes", "WAA_标准模板.xlsx")


def check(cond, label):
    print(("PASS " if cond else "FAIL ") + label)
    return bool(cond)


class FakeWriter:
    def __init__(self, row=7):
        self.row = row
        self.frames = []

    def add_axis_row(self, frame):
        self.frames.append(frame)
        return self.row


class FakeReporter:
    def __init__(self, row=8):
        self.row = row
        self.frames = []

    def add_axis_row(self, frame):
        self.frames.append(frame)
        return self.row


def fake_executor():
    ex = Executor.__new__(Executor)
    ex.counter = SimpleNamespace(frame=100)
    ex.start_offset = 0
    ex.reporter = FakeReporter()
    ex.mark_writer = None
    ex.mark_writer_error = ""
    ex.mark_suspended = False
    ex.axis = SimpleNamespace(actions=[], settings={})
    ex.axis_running = False
    ex.current_action = None
    return ex


def main() -> int:
    os.makedirs(TMP, exist_ok=True)
    checks = []

    # 1) offset / frame
    ex = fake_executor()
    ex.set_offset(delta=1)
    checks.append(check(ex.current_frame() == 101, "offset +1"))
    ex.set_frame(500)
    checks.append(check(ex.current_frame() == 500, "frame=500"))
    checks.append(check(ex.start_offset == 400, f"offset={ex.start_offset}"))

    # 2) 没有 write：F5 走主轴 reporter，并追加无操作行
    row = ex.mark_frame()
    checks.append(check(row == 8 and ex.reporter.frames == [500],
                        f"fallback F5 row={row}"))
    checks.append(check(len(ex.axis.actions) == 1,
                        "fallback append axis action"))

    # 3) 有独立 write：F5 只写 write，不动主轴
    ex2 = fake_executor()
    writer = FakeWriter()
    ex2.mark_writer = writer
    row2 = ex2.mark_frame()
    checks.append(check(row2 == 7 and writer.frames == [100],
                        f"write F5 row={row2}"))
    checks.append(check(len(ex2.axis.actions) == 0,
                        "write 不动主轴 actions"))

    # 4) read 期间拒绝 F5
    ex3 = fake_executor()
    ex3.mark_suspended = True
    checks.append(check(ex3.mark_frame() is None, "read 期间 F5 拒绝"))

    # 5) write 文件损坏/不存在：报错，不回退
    ex4 = fake_executor()
    ex4.mark_writer_error = "write 文件不存在"
    checks.append(check(ex4.mark_frame() is None
                        and not ex4.reporter.frames,
                        "write 错误不回退"))

    # 6) MarkWriter 真正追加一行
    dst = os.path.join(TMP, "mark_test.xlsx")
    if os.path.exists(dst):
        os.remove(dst)
    shutil.copy2(SRC, dst)
    mw = MarkWriter(dst)
    r = mw.add_axis_row(123)
    mw.close()
    wb = load_workbook(dst)
    ws = wb["作战记录"]
    checks.append(check(ws.cell(r, 2).value == 123,
                        f"MarkWriter 第 {r} 行 = {ws.cell(r, 2).value}"))
    wb.close()

    # 7) showframe 透明度解析：百分比和小数都支持
    checks.append(check(_parse_alpha("55") == 0.55, "透明度 55 -> 0.55"))
    checks.append(check(_parse_alpha("0.55") == 0.55, "透明度 0.55 -> 0.55"))
    checks.append(check(_parse_alpha("100") == 1.0, "透明度 100 -> 1.0"))
    checks.append(check(_parse_alpha("5") is None, "透明度 5 被拒绝"))
    checks.append(check(_parse_alpha("abc") is None, "透明度 abc 被拒绝"))

    # 8) 状态文件同时带帧号和 alpha；窗口子进程只认这个文件
    sf = ShowFrameWindow(lambda: ("123", "00:04:03", "2/30"), alpha=0.55)
    sf.state_path = os.path.join(TMP, "showframe_state.json")
    sf.set_alpha(0.35)
    with open(sf.state_path, "r", encoding="utf-8") as f:
        state = json.load(f)
    checks.append(check(state["frame"] == "123" and state["alpha"] == 0.35,
                        f"showframe 状态 = {state}"))

    # 9) 没 read 轴时的预览计数器
    preview = _make_preview_counter()
    checks.append(check(preview.rate_percent == 100.0
                        and preview.counter.rate_candidates
                        == [50.0, 100.0, 110.0],
                        f"预览候选 = {preview.counter.rate_candidates}"))
    frames, _ = build(rate=100.0, n=120)
    for img in frames:
        preview.update(cv2.cvtColor(img, cv2.COLOR_GRAY2BGR))
    checks.append(check(preview.armed and preview.frame == len(frames) - 1,
                        f"预览计数 = {preview.frame}/{len(frames) - 1}"))
    blank = np.full((1080, 1920, 3), 50, np.uint8)
    for _ in range(40):
        preview.update(blank)
    checks.append(check(preview._out_of_battle and not preview.in_battle,
                        "费用条消失 40 帧 -> 离开战斗、停止检测"))
    frames2, _ = build(rate=100.0, n=60)
    for img in frames2:
        preview.update(cv2.cvtColor(img, cv2.COLOR_GRAY2BGR))
    checks.append(check(preview.frame == len(frames2) - 1 and preview.in_battle,
                        f"新关卡归零后计数 = {preview.frame}/{len(frames2) - 1}"))

    # 10) status 读取运行信息：帧号/时间/费用条/回费/倍速/状态
    axis = SimpleNamespace(path=os.path.join(TMP, "fake.xlsx"),
                           settings={}, actions=[])
    rep = ExcelReporter(axis, {"base": 100.0},
                        os.path.join(TMP, "reports"))
    last = SimpleNamespace(neg=False, phase=11, cycle_len=31, m=2.0,
                           bullet=False, running=True)
    snap_counter = SimpleNamespace(
        frame=630, last=last,
        counter=SimpleNamespace(rate_percent=100.0, manual_rate=None,
                                neg=False, spec_error=""))
    snap_ex = SimpleNamespace(phase="运行中", stream=None)
    rep.bind_runtime(snap_counter, snap_ex)
    data = rep.status_snapshot()
    checks.append(check(data["frame"] == "630" and data["time"] == "00:21:00",
                        f"status 帧/时间 = {data['frame']} {data['time']}"))
    checks.append(check(data["cost_frame"] == "11/31",
                        f"status 费用条 = {data['cost_frame']}"))
    checks.append(check(data["rate"] == "100%" and data["speed"] == "2X",
                        f"status 回费/倍速 = {data['rate']} {data['speed']}"))
    checks.append(check(data["status"] == "运行中",
                        f"status 状态 = {data['status']}"))

    class FakeOCR:
        def __init__(self):
            self.submitted = 0
        def submit(self, img):
            self.submitted += 1
        def values(self):
            return {"cost": 12}

    rep._ocr = FakeOCR()
    snap_counter.in_battle = False
    snap_ex.stream = SimpleNamespace(latest=lambda: object())
    data = rep.status_snapshot()
    checks.append(check(rep._ocr.submitted == 0 and data["cost"] == "",
                        "不在关卡 -> 不做 OCR"))
    snap_counter.in_battle = True
    data = rep.status_snapshot()
    checks.append(check(rep._ocr.submitted == 1 and data["cost"] == "12",
                        "在关卡 -> 做 OCR"))

    # 11) 中途开始的部署栈重建 + 已过帧回费事件
    actions = [
        AxisAction(frame=10, kind="部署", oper="A"),
        AxisAction(frame=20, kind="部署", oper="B"),
        AxisAction(frame=30, kind="撤退", oper="A"),
        AxisAction(frame=100, kind="部署", oper="C"),
    ]
    mid = Executor.__new__(Executor)
    mid.axis = SimpleNamespace(actions=actions)
    mid.deployed_stack = []
    mid._rebuild_stack_for_start(60, 0, actions)
    checks.append(check([r["name"] for r in mid.deployed_stack] == ["B"],
                        f"中途部署栈 = {[r['name'] for r in mid.deployed_stack]}"))
    kept, last = Executor._trim_rate_events(
        [(0, 80.0, actions[0]), (50, None, None),
         (100, 50.0, None), (200, None, None)], 200, 120)
    checks.append(check(kept == [(120, 50.0, None), (200, None, None)]
                        and last == 200,
                        f"回费事件裁剪 = {kept}, last={last}"))

    class SkipReporter:
        def __init__(self):
            self.calls = []
        def action_done(self, *a, **k):
            self.calls.append(a)
        def set_error(self, *a, **k):
            pass

    past_ex = Executor.__new__(Executor)
    past_ex.axis = SimpleNamespace(actions=[
        AxisAction(frame=10, kind="部署", oper="A"),
        AxisAction(frame=20, kind="技能", oper="A"),
    ])
    past_ex.counter = SimpleNamespace(frame=200)
    past_ex.start_offset = 0
    past_ex.deployed_stack = []
    past_ex.stop_requested = threading.Event()
    past_ex.manual_pause = threading.Event()
    past_ex.end_pause = False
    past_ex.phase = "准备中"
    past_ex.axis_running = False
    past_ex.reporter = SkipReporter()
    past_ex._write_pause_log = lambda: None
    past_ex.resume = lambda: None
    past_ex.run(max_actions=2)
    checks.append(check(not past_ex.reporter.calls,
                        "已过帧动作不写运行结果"))

    # 12) 没 read 时的 offset/frame：offset 永远绝对，frame 支持 ±
    sess = ConsoleSession.__new__(ConsoleSession)
    sess.counter = SimpleNamespace(frame=100)
    sess.executor = None
    sess.preview_offset = 5
    sess.offset_command("+5")
    checks.append(check(sess.preview_offset == 5, "预览 offset +5 = 5"))
    sess.offset_command("-3")
    checks.append(check(sess.preview_offset == -3, "预览 offset -3 = -3"))
    sess.frame_command("+10")
    checks.append(check(sess.preview_offset == 7, "预览 frame +10"))
    sess.frame_command("200")
    checks.append(check(sess.preview_offset == 100, "预览 frame 200"))
    sess.frame_command("-5")
    checks.append(check(sess.preview_offset == 95, "预览 frame -5"))

    # 13) 退出关卡自动 stop 的判定逻辑
    from waa.core.waa_counter import RESET_AFTER_OUT_OF_BATTLE_S

    def bare_session():
        s = ConsoleSession.__new__(ConsoleSession)
        s._battle_seen = False
        return s

    # 13a) 没进过战斗（in_battle 一直是 False）：HUD 丢再久也不算退出
    s = bare_session()
    s.counter = SimpleNamespace(in_battle=False,
                                out_of_battle_s=lambda: 99.0)
    checks.append(check(not s._battle_exit_should_stop(),
                        "没进过战斗 -> 不触发自动 stop"))

    # 13b) 进过战斗：HUD 短暂丢失（<3s）不触发
    s.counter.in_battle = True
    checks.append(check(s._battle_exit_should_stop() is False
                        and s._battle_seen,
                        "in_battle=True -> 记录已见战斗"))
    s.counter.in_battle = False
    s.counter.out_of_battle_s = lambda: 1.0
    checks.append(check(not s._battle_exit_should_stop(),
                        "HUD 丢失 1s -> 不触发"))

    # 13c) HUD 消失 ≥3s（和帧号归零同一个阈值）-> 触发
    s.counter.out_of_battle_s = lambda: RESET_AFTER_OUT_OF_BATTLE_S
    checks.append(check(s._battle_exit_should_stop(),
                        f"HUD 丢失 {RESET_AFTER_OUT_OF_BATTLE_S}s -> 触发"))

    # 13d) 真计数器：离开战斗后 out_of_battle_s() 随时间增长，回战斗归零
    frames3, _ = build(rate=100.0, n=40)
    for img in frames3:
        preview.update(cv2.cvtColor(img, cv2.COLOR_GRAY2BGR))
    checks.append(check(preview.out_of_battle_s() == 0.0,
                        "在战斗 -> out_of_battle_s=0"))
    for _ in range(40):
        preview.update(blank)
    preview._out_of_battle_since = time.time() - 3.5
    dur = preview.out_of_battle_s()
    checks.append(check(3.4 <= dur < 4.0,
                        f"离开战斗 -> out_of_battle_s={dur:.1f}"))

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
