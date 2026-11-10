"""原地写入离线测试：不跑游戏，验证模板 -> 填轴 -> 直接写回同一个 xlsx。"""

from __future__ import annotations

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
from waa.core.excel_report import ExcelReporter, _frame_time_text  # noqa: E402

TMP = os.path.join(ROOT, "out", "_inplace_test")
SRC = os.path.join(ROOT, "axes", "WAA_标准模板.xlsx")
DST = os.path.join(TMP, "WAA_inplace_test.xlsx")


class FakeLast:
    neg = False
    running = True
    phase = 11
    cycle_len = 30
    m = 1.0
    bullet = False


class FakeCounter:
    spec_error = ""


class FakeRuntime:
    frame = 630
    frame_hi = 630
    rate_percent = 100.0
    last = FakeLast()
    counter = FakeCounter()


class FakeExecutor:
    manual_pause = threading.Event()


def prepare() -> None:
    os.makedirs(TMP, exist_ok=True)
    for path in glob.glob(os.path.join(TMP, "WAA_inplace_test*")):
        try:
            os.remove(path)
        except OSError:
            pass
    shutil.copy2(SRC, DST)
    wb = load_workbook(DST)
    ws = wb["作战记录"]
    # 测试自带地图代号，不依赖用户手动改过的模板设置。
    ws["M2"], ws["N2"] = "关卡代号", "1-7"
    ws["B2"] = 150
    ws["D2"] = "部署"
    ws["E2"] = "斑点"
    ws["F2"] = "D2"
    ws["G2"] = "右"
    ws["I2"] = "test"
    ws["J2"] = True
    # 故意空出第 3 行，验证结果按实际 Excel 行号写，不会错位。
    ws["B4"] = 240
    ws["D4"] = "部署"
    ws["E4"] = "Castle-3"
    ws["F4"] = "D5"
    ws["G4"] = "下"
    ws["H4"] = "近卫车"
    ws["J4"] = True
    # 模拟上一局留下的结果：新一局必须清掉，尤其是空行 K3 的旧值。
    ws["K2"] = "上一局结果"
    ws["K3"] = "上一局残留"
    ws["K4"] = "上一局结果"
    wb.save(DST)
    wb.close()


def check(cond: bool, label: str) -> bool:
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def main() -> int:
    prepare()
    axis = read_axis(DST)
    print("axis:", axis.settings, axis.map_code, len(axis.actions),
          [a.excel_row for a in axis.actions])
    if len(axis.actions) != 2 or [a.excel_row for a in axis.actions] != [2, 4]:
        print("FAIL 轴没有读到第 2、4 行两条动作")
        return 1

    rep = ExcelReporter(
        axis,
        {"base": 100.0, "declared": [100.0]},
        os.path.join(TMP, "reports"),
        enabled=True,
        quit_on_stop=True,
        in_place=True,
        backup=False,
    )
    rep.bind_runtime(FakeRuntime(), FakeExecutor())
    ok = rep.start(wait_ready=True)
    print("start:", ok, "report:", rep.report_path)
    if not ok:
        return 1
    save_ok = rep.save_workbook()
    print("save_workbook:", save_ok)

    for i, action in enumerate(axis.actions):
        rep.action_start(i, action)
        time.sleep(0.25)
        rep.action_done(i, action, action.frame, action.frame, "成功")
        time.sleep(0.35)
    rep.finish(True)
    time.sleep(0.3)
    rep.stop()

    checks: list[bool] = []
    checks.append(check(_frame_time_text(71) == "00:02:11",
                        f"时间格式 71 = {_frame_time_text(71)}"))
    checks.append(check(_frame_time_text(1799) == "00:59:29",
                        f"时间格式 1799 = {_frame_time_text(1799)}"))
    checks.append(check(_frame_time_text(1800) == "01:00:00",
                        f"时间格式 1800 = {_frame_time_text(1800)}"))
    checks.append(check(save_ok, "save_workbook"))
    checks.append(check(os.path.abspath(rep.report_path) == os.path.abspath(DST),
                        "报告路径就是原文件"))

    wb = load_workbook(DST)
    ws = wb["作战记录"]
    mid = wb["运行信息"]
    map_ws = wb["地图示意"]

    checks.append(check(wb.sheetnames == ["作战记录", "运行信息", "地图示意"],
                        f"三页顺序 {wb.sheetnames}"))
    checks.append(check([ws.cell(row=2, column=c).value for c in range(2, 11)]
                        == [150, None, "部署", "斑点", "D2", "右", None, "test", True],
                        "第 2 行动作没被改动"))
    checks.append(check(all(ws.cell(row=3, column=c).value in (None, "")
                            for c in range(2, 11)),
                        "空着的第 3 行保持空白"))
    checks.append(check([ws.cell(row=4, column=c).value for c in range(2, 11)]
                        == [240, None, "部署", "Castle-3", "D5", "下", "近卫车", None, True],
                        "第 4 行动作没被改动"))
    checks.append(check(ws["K2"].value is True, f"结果 K2 = {ws['K2'].value!r}"))
    checks.append(check(ws["K3"].value in (None, ""),
                        f"空行 K3 没被写 = {ws['K3'].value}"))
    checks.append(check(ws["K4"].value is True, f"结果 K4 = {ws['K4'].value!r}"))
    checks.append(check((ws["A2"].value, ws["A4"].value) == (1, 2),
                        f"当前执行顺序 = {ws['A2'].value},{ws['A4'].value}"))
    checks.append(check(str(mid["B1"].value).startswith("轴完成"),
                        f"状态 = {mid['B1'].value}"))
    checks.append(check(mid["A5"].value == "当前帧号",
                        f"标签 A5 = {mid['A5'].value}"))
    checks.append(check(str(mid["B5"].value) == "630", f"当前帧 = {mid['B5'].value}"))
    checks.append(check(mid["A6"].value == "当前时间",
                        f"标签 A6 = {mid['A6'].value}"))
    checks.append(check(mid["B6"].value == "00:21:00", f"时间 = {mid['B6'].value}"))
    checks.append(check(mid["A7"].value == "费用条",
                        f"标签 A7 = {mid['A7'].value}"))
    checks.append(check(mid["B7"].value == "11/30", f"费用条 = {mid['B7'].value}"))
    checks.append(check(mid["B8"].value == "100%", f"回费速度 = {mid['B8'].value}"))

    d2 = map_ws.cell(row=6, column=3)
    d5 = map_ws.cell(row=6, column=6)
    checks.append(check(d2.value == "斑点→", f"D2 地图 = {d2.value}"))
    checks.append(check(str(d2.fill.fgColor.rgb).endswith("9DC3E6"),
                        f"D2 颜色 = {d2.fill.fgColor.rgb}"))
    checks.append(check(d5.value == "近卫车↓", f"D5 地图 = {d5.value}"))
    checks.append(check(str(d5.fill.fgColor.rgb).endswith("9DC3E6"),
                        f"D5 颜色 = {d5.fill.fgColor.rgb}"))

    wb.close()

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
