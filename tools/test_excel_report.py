"""Excel 运行面板离线测试：不跑游戏，只走 action_start/done/finish。"""

from __future__ import annotations

import os
import shutil
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from openpyxl import load_workbook  # noqa: E402

from waa.core.axis import read_axis  # noqa: E402
from waa.core.excel_report import ExcelReporter  # noqa: E402

TMP = os.path.join(ROOT, "out", "_report_test")
SRC = os.path.join(ROOT, "axes", "WAA_sample_1-7.xlsx")
SRC_FRAME = os.path.join(TMP, "WAA_sample_帧号.xlsx")


def prepare_axis() -> str:
    """标准模板只认“帧号”；旧 sample 是“帧数”，这里复制后改表头。"""
    os.makedirs(TMP, exist_ok=True)
    shutil.copy2(SRC, SRC_FRAME)
    wb = load_workbook(SRC_FRAME)
    wb["作战记录"]["B1"] = "帧号"
    wb.save(SRC_FRAME)
    wb.close()
    return SRC_FRAME


def main() -> int:
    axis = read_axis(prepare_axis())
    rep = ExcelReporter(axis, {"base": 100.0, "declared": [100.0]},
                        os.path.join(ROOT, "out", "_report_test"),
                        enabled=True, quit_on_stop=True)
    rep.start(wait_ready=True)
    if not rep.report_path:
        print("面板没启动")
        return 1
    time.sleep(0.4)

    acts = axis.actions
    rep.action_start(0, acts[0])
    time.sleep(0.3)
    rep.action_done(0, acts[0], 150, 150)
    if len(acts) >= 2:
        rep.action_start(1, acts[1])
        time.sleep(0.2)
        rep.action_done(1, acts[1], 240, 240)
    if len(acts) >= 3:
        rep.action_start(2, acts[2])
        time.sleep(0.2)
        rep.action_done(2, acts[2], 240, 240)
    rep.finish(True)
    time.sleep(0.5)
    rep.stop()

    print("report:", rep.report_path)
    wb = load_workbook(rep.report_path)
    print("sheets:", wb.sheetnames)
    mid = wb["运行信息"]
    print("第2页状态:", mid["A1"].value, mid["B1"].value, "|",
          mid["A2"].value, mid["B2"].value, "|",
          mid["A3"].value, mid["B3"].value, "|",
          mid["A5"].value, mid["B5"].value, "|",
          mid["A6"].value, mid["B6"].value)
    rec = wb["作战记录"]
    print("作战记录 K2:", rec["K2"].value, " K3:", rec["K3"].value,
          " A2:", rec["A2"].value, " 设置:", rec["L2"].value, rec["M2"].value,
          rec["L3"].value, rec["M3"].value, rec["N3"].value)
    map_ws = wb["地图示意"]
    print("地图 D2 格:", map_ws.cell(row=6, column=3).value,
          " D5 格:", map_ws.cell(row=6, column=6).value)
    print("sheets:", wb.sheetnames)
    return 0


if __name__ == "__main__":
    sys.exit(main())
