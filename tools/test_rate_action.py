"""离线测试：回费速度行的解析、事件展开和计数器切换。"""

from __future__ import annotations

import os
import shutil
import sys
import threading
from types import SimpleNamespace

import cv2

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from openpyxl import load_workbook  # noqa: E402

from waa.core.axis import AxisAction, parse_rate_value, read_axis  # noqa: E402
from waa.core.cost_model import (  # noqa: E402
    NEG_SPEED, SCALE, SPEED, measured_width, raw_width,
)
from waa.core.excel_report import ExcelReporter  # noqa: E402
from waa.core.executor import Executor  # noqa: E402
from waa.core.waa_counter import WAACounter  # noqa: E402
from synth_neg_test import render  # noqa: E402

TMP = os.path.join(ROOT, "out", "_rate_test")
SRC = os.path.join(ROOT, "axes", "WAA_标准模板.xlsx")
DST = os.path.join(TMP, "WAA_rate_test.xlsx")


def check(cond, label):
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def test_parse_value():
    checks = []
    for raw, want, label in (
        ("自动", "auto", "自动"),
        ("auto", "auto", "auto"),
        ("80%", 80.0, "80%"),
        ("0.8", 80.0, "0.8"),
        ("80", 80.0, "80"),
        ("100%", 100.0, "100%"),
        ("", None, "空白"),
        ("abc", None, "非法"),
    ):
        got = parse_rate_value(raw)
        checks.append(check(got == want, f"parse {label}: {raw!r} -> {got}"))
    return checks


def test_read_axis():
    os.makedirs(TMP, exist_ok=True)
    shutil.copy2(SRC, DST)
    wb = load_workbook(DST)
    ws = wb["作战记录"]
    ws["B2"] = 100
    ws["C2"] = 200
    ws["D2"] = "回费速度"
    ws["E2"] = "80%"
    ws["B3"] = 300
    ws["D3"] = "回费速度"
    ws["E3"] = "自动"
    wb.save(DST)
    wb.close()

    axis = read_axis(DST)
    checks = []
    checks.append(check(len(axis.actions) == 2,
                        f"读入两条回费速度: {len(axis.actions)}"))
    a, b = axis.actions
    checks.append(check(a.kind == "回费速度" and a.rate_percent == 80.0,
                        f"80% 行: kind={a.kind} rate={a.rate_percent}"))
    checks.append(check(a.frame == 100 and a.frame2 == 200,
                        f"80% 区间: {a.frame}~{a.frame2}"))
    checks.append(check(b.rate_auto is True and b.rate_percent is None,
                        f"自动行: auto={b.rate_auto} rate={b.rate_percent}"))
    return checks


def test_build_events():
    a = AxisAction(frame=100, frame2=200, kind="回费速度",
                   rate_percent=80.0, excel_row=2)
    b = AxisAction(frame=300, kind="回费速度",
                   rate_auto=True, excel_row=3)
    a._index = 0
    b._index = 1
    events, last = Executor._build_rate_events([a, b])
    checks = []
    checks.append(check(last == 300, f"最后事件帧: {last}"))
    checks.append(check(events == [
        (100, 80.0, a), (201, None, None), (300, None, b)],
        f"事件序列: {events}"))
    return checks


def test_counter_switch():
    action = AxisAction(frame=100, frame2=200, kind="回费速度",
                        rate_percent=80.0, excel_row=2)
    action._index = 0
    called = []
    afc = WAACounter(100.0)
    afc.rate_callback = called.append
    afc.set_rate_events([(100, 80.0, action), (201, None, None)])
    checks = []

    afc.counter.frame = 100
    with afc._rate_lock:
        afc._apply_rate_events_locked()
    checks.append(check(afc.counter.manual_rate == 80.0,
                        f"80% 生效: manual={afc.counter.manual_rate}"))
    checks.append(check(afc.counter.required == SCALE,
                        f"80% required 固定: {afc.counter.required}"))
    checks.append(check(afc.counter.eps == 167,
                        f"80% eps 固定: {afc.counter.eps}"))
    checks.append(check(afc.counter.normal_speed == 447392,
                        f"80% speed: {afc.counter.normal_speed}"))
    checks.append(check(afc.counter.neg_speed == 223696,
                        f"手动负费 = 半速: {afc.counter.neg_speed}"))
    checks.append(check(called == [action], f"动作回调: {called}"))

    afc.counter.frame = 201
    with afc._rate_lock:
        afc._apply_rate_events_locked()
    checks.append(check(afc.counter.manual_rate is None,
                        f"结束后恢复默认: manual={afc.counter.manual_rate}"))
    checks.append(check(afc.counter.required == SCALE
                        and afc.counter.eps == 167,
                        f"默认 required/eps: {afc.counter.required}/"
                        f"{afc.counter.eps}"))
    checks.append(check(afc.counter.normal_speed == SPEED
                        and afc.counter.neg_speed == NEG_SPEED,
                        f"默认 speed/负费: {afc.counter.normal_speed}/"
                        f"{afc.counter.neg_speed}"))
    checks.append(check(afc.rate_events_done(), "事件已全部处理"))
    return checks


def test_transition():
    """合成一条 100% -> 80% 的费用条，确认切换帧不丢帧。"""
    dt = 69905 / 2 ** 21
    c = 1.0
    frames = []
    switch = 30
    for i in range(120):
        rate = 80.0 if i >= switch else 100.0
        w = measured_width(raw_width(1.0 - c))
        nxt = c - dt * (rate / 100.0)
        if nxt < 1e-5:
            nxt += 1.0
        nw = measured_width(raw_width(1.0 - nxt))
        cap = (w == 0 and nw > 0)
        frames.append(cv2.cvtColor(render(w, cap, False),
                                   cv2.COLOR_GRAY2BGR))
        c = nxt

    action = AxisAction(frame=switch, kind="回费速度", rate_percent=80.0)
    action._index = 0
    afc = WAACounter(100.0)
    afc.set_rate_events([(switch, 80.0, action)])
    afc.arm(0)
    warns = []
    for img in frames:
        afc.update(img)
        if afc.counter.warn:
            warns.append((afc.frame, afc.counter.warn))

    checks = [
        check(afc.frame == len(frames) - 1
              or afc.frame_hi == len(frames) - 1,
              f"切换后帧号: {afc.frame}~{afc.frame_hi} / {len(frames) - 1}"),
        check(afc.counter.manual_rate == 80.0,
              f"切换后手速: {afc.counter.manual_rate}"),
        check(not warns, f"切换过程无匹配失败: {warns[:3]}"),
    ]
    return checks


def test_result_recording():
    action = AxisAction(frame=100, frame2=200, kind="回费速度",
                        rate_percent=80.0, excel_row=2)
    action._index = 3
    checks = []

    # 回调：无主循环动作时，只写结果，不改当前动作显示。
    class FakeRep:
        def __init__(self):
            self.calls = []

        def record_simple_action(self, idx, act, frame, note):
            self.calls.append((idx, act, frame, note))

    ex = Executor.__new__(Executor)
    ex.reporter = FakeRep()
    Executor._rate_action_applied(ex, action)
    checks.append(check(ex.reporter.calls
                        == [(3, action, 100, "成功")],
                        f"回费动作回调: {ex.reporter.calls}"))

    # ExcelReporter.record_simple_action：结果布尔 True，并分配顺序。
    rep = ExcelReporter.__new__(ExcelReporter)
    rep._lock = threading.Lock()
    rep._order = 0
    rep._results = {}
    rep._result_dirty = False
    rep._dirty = False
    rep.record_simple_action(3, action, 100, "成功")
    r = rep._results.get(3, {})
    checks.append(check(r.get("ok") is True and r.get("order") == 1,
                        f"回费结果: {r}"))
    return checks


def test_display():
    rep = ExcelReporter.__new__(ExcelReporter)
    rep.rate_info = {"base": 100.0, "declared": [100.0, 50.0]}
    rep._counter = SimpleNamespace(
        counter=SimpleNamespace(manual_rate=80.0))
    checks = [
        check(rep._realtime_rate_text(False) == "80%（手动）",
              f"手动显示: {rep._realtime_rate_text(False)}"),
        check(rep._realtime_rate_text(True) == "40%（手动，欠费）",
              f"手负费显示: {rep._realtime_rate_text(True)}"),
    ]
    rep._counter.counter.manual_rate = None
    checks.append(check(rep._realtime_rate_text(False) == "100%",
                        f"默认显示: {rep._realtime_rate_text(False)}"))
    return checks


def main() -> int:
    checks = []
    checks += test_parse_value()
    checks += test_read_axis()
    checks += test_build_events()
    checks += test_counter_switch()
    checks += test_transition()
    checks += test_result_recording()
    checks += test_display()
    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
