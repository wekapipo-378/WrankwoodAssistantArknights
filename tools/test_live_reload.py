"""离线测试：F6 暂停编辑后，重载未执行动作。"""

from __future__ import annotations

import os
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from openpyxl import load_workbook  # noqa: E402

from waa.core.axis import read_axis  # noqa: E402
from waa.core.executor import Executor  # noqa: E402

TMP = os.path.join(ROOT, "out", "_reload_test")
SRC = os.path.join(ROOT, "axes", "WAA_标准模板.xlsx")
DST = os.path.join(TMP, "WAA_reload_test.xlsx")


class FakeReporter:
    enabled = True
    started = True
    in_place = True

    def __init__(self, path):
        self.report_path = path
        self.saved = 0

    def save_workbook(self, timeout=5.0):
        self.saved += 1
        return True


def check(cond, label):
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
    ws["D2"] = "部署"
    ws["E2"] = "斑点"
    ws["F2"] = "D2"
    ws["G2"] = "右"
    ws["B4"] = 200
    ws["D4"] = "部署"
    ws["E4"] = "Castle-3"
    ws["F4"] = "D5"
    ws["G4"] = "下"
    ws["H4"] = "近卫车"
    wb.save(DST)
    wb.close()

    axis = read_axis(DST)
    print("axis actions:", [(a.frame, a.kind, a.oper, a.excel_row)
                            for a in axis.actions])
    if len(axis.actions) != 2:
        print("FAIL 初始动作数不是 2")
        return 1

    rep = FakeReporter(DST)
    ex = Executor.__new__(Executor)
    ex.axis = axis
    ex.reporter = rep
    ex.pending_from = 0
    ex.current_frame = lambda: 0

    # 1) 改两行的帧，重载后都应更新。
    wb = load_workbook(DST)
    ws = wb["作战记录"]
    ws["B2"] = 110
    ws["B4"] = 210
    wb.save(DST)
    wb.close()
    ok = Executor.reload_pending_actions(ex)
    checks = [check(ok, "reload returned True"),
              check(axis.actions[0].frame == 110,
                    f"row2 frame = {axis.actions[0].frame}"),
              check(axis.actions[1].frame == 210,
                    f"row4 frame = {axis.actions[1].frame}"),
              check(rep.saved == 1, f"saved = {rep.saved}")]

    # 2) pending_from=1：只更新第 2 个动作（Excel 第 4 行）。
    wb = load_workbook(DST)
    ws = wb["作战记录"]
    ws["B2"] = 120
    ws["B4"] = 220
    wb.save(DST)
    wb.close()
    ex.pending_from = 1
    ok2 = Executor.reload_pending_actions(ex)
    checks.append(check(ok2, "reload2 returned True"))
    checks.append(check(axis.actions[0].frame == 110,
                        f"past row2 stays = {axis.actions[0].frame}"))
    checks.append(check(axis.actions[1].frame == 220,
                        f"pending row4 = {axis.actions[1].frame}"))

    # 3) 运行时把未执行行的“是否运行”改成 False，重载后生效。
    wb = load_workbook(DST)
    ws = wb["作战记录"]
    ws["J4"] = False
    wb.save(DST)
    wb.close()
    ok_enabled = Executor.reload_pending_actions(ex)
    checks.append(check(ok_enabled and axis.actions[1].enabled is False,
                        f"是否运行重载 = {axis.actions[1].enabled}"))

    # 4) 把行 4 的操作清空：读不出，保留旧值，不能崩。
    wb = load_workbook(DST)
    ws = wb["作战记录"]
    ws["C4"] = None
    wb.save(DST)
    wb.close()
    ok3 = Executor.reload_pending_actions(ex)
    checks.append(check(axis.actions[1].frame == 220,
                        f"invalid row keeps old = {axis.actions[1].frame}"))

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
