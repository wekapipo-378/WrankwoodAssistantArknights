"""地图实时刷新离线测试：动作开始/结束、负费状态在面板运行中是否可见。"""

from __future__ import annotations

import atexit
import glob
import os
import shutil
import sys
import threading
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from openpyxl import load_workbook  # noqa: E402

from waa.core.axis import read_axis  # noqa: E402
from waa.core.excel_report import ExcelReporter  # noqa: E402

TMP = os.path.join(ROOT, "out", "_map_live_test")
SRC = os.path.join(ROOT, "axes", "WAA_标准模板.xlsx")
DST = os.path.join(TMP, "WAA_map_live.xlsx")


class FakeLast:
    neg = False
    running = True


class FakeCounter:
    spec_error = ""
    neg = False


class FakeRuntime:
    frame = 0
    frame_hi = 0
    rate_percent = 100.0
    last = FakeLast()
    counter = FakeCounter()


class FakeExecutor:
    manual_pause = threading.Event()


def prepare() -> None:
    os.makedirs(TMP, exist_ok=True)
    for path in glob.glob(os.path.join(TMP, "WAA_map_live*")):
        try:
            os.remove(path)
        except OSError:
            pass
    shutil.copy2(SRC, DST)
    wb = load_workbook(DST)
    ws = wb["作战记录"]
    # 测试自带地图代号，不依赖用户手动改过的模板设置。
    ws["M2"], ws["N2"] = "关卡代号", "1-7"
    # 新轴：B=帧号1、C=帧号2、D=操作、E=干员、F=坐标、G=朝向、H=简称。
    ws["B2"], ws["D2"], ws["E2"], ws["F2"], ws["G2"] = 30, "部署", "斑点", "D2", "右"
    ws["B4"], ws["D4"], ws["E4"], ws["F4"], ws["G4"], ws["H4"] = (
        150, "部署", "Castle-3", "D5", "下", "近卫车")
    wb.save(DST)
    wb.close()


def cells() -> tuple:
    wb = load_workbook(DST)
    mid = wb["运行信息"]
    map_ws = wb["地图示意"]
    d2 = map_ws.cell(row=6, column=3)
    d5 = map_ws.cell(row=6, column=6)
    out = (d2.value, str(d2.fill.fgColor.rgb), d5.value,
           str(d5.fill.fgColor.rgb), mid["B8"].value, mid["B5"].value)
    wb.close()
    return out


def main() -> int:
    prepare()
    axis = read_axis(DST)
    fake = FakeRuntime()
    rep = ExcelReporter(
        axis, {"base": 100.0, "declared": [100.0, 50.0]},
        os.path.join(TMP, "reports"),
        enabled=True, quit_on_stop=True, in_place=True, backup=False,
    )
    rep.bind_runtime(fake, FakeExecutor())
    if not rep.start(wait_ready=True):
        return 1
    atexit.register(rep.stop)   # 测试中途抛异常也关掉 Excel，别留锁
    print("start", cells())
    checks: list[bool] = []

    def check(cond, label):
        print(("PASS " if cond else "FAIL ") + label)
        checks.append(cond)

    check([a.frame for a in axis.actions] == [30, 150],
          f"“帧号”表头读帧 = {[a.frame for a in axis.actions]}")
    # bring_to_front 现在返回真实结果（能不能真抢到前台）；后台测试进程可能
    # 抢不到前台，这里只验证它正常返回布尔、不抛异常，不强求 True。
    front = rep.bring_to_front()
    check(isinstance(front, bool), f"bring_to_front 返回布尔 = {front}")

    fake.frame = fake.frame_hi = 30
    rep.action_start(0, axis.actions[0])
    time.sleep(0.4)
    rep.save_workbook()
    a = cells()
    check(a[0] == "斑点→" and "FFD966" in a[1],
          f"动作开始 D2 黄色 = {a}")

    fake.frame = fake.frame_hi = 31
    rep.action_done(0, axis.actions[0], 30, 31, "成功")
    time.sleep(0.4)
    rep.save_workbook()
    b = cells()
    check(b[0] == "斑点→" and "9DC3E6" in b[1],
          f"部署完成 D2 蓝色 = {b}")

    fake.frame = fake.frame_hi = 150
    rep.action_start(1, axis.actions[1])
    time.sleep(0.4)
    rep.save_workbook()
    c = cells()
    check(c[2] == "近卫车↓" and "FFD966" in c[3],
          f"动作开始 D5 黄色 = {c}")
    check("9DC3E6" in c[1], f"动作开始 D2 仍蓝 = {c}")

    fake.frame = fake.frame_hi = 151
    rep.action_done(1, axis.actions[1], 150, 151, "成功")
    time.sleep(0.4)
    rep.save_workbook()
    d = cells()
    check(d[2] == "近卫车↓" and "9DC3E6" in d[3],
          f"部署完成 D5 蓝色 = {d}")

    # 负费实时速度
    fake.last.neg = True
    fake.counter.neg = True
    time.sleep(0.5)
    rep.save_workbook()
    e = cells()
    check(e[4] == "50%（欠费）", f"地图页负费速度 = {e[4]}")
    fake.last.neg = False
    fake.counter.neg = False
    time.sleep(0.5)
    rep.save_workbook()
    f = cells()
    check(f[4] == "100%", f"地图页正常速度 = {f[4]}")

    # 失败动作只写“失败”，不写原因。
    rep.action_done(0, axis.actions[0], 30, None, "失败（部署未成功）")
    time.sleep(0.4)
    rep.save_workbook()
    wb = load_workbook(DST)
    result = wb["作战记录"]["K2"].value
    wb.close()
    check(result is False, f"失败结果布尔 = {result!r}")

    rep.finish(True)
    time.sleep(0.3)
    rep.stop()
    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
