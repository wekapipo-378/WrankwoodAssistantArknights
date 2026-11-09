"""离线测试：分组排序、行映射、公式引用重写。"""

from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa.core.axis import AxisAction  # noqa: E402
from waa.core.excel_sort import (  # noqa: E402
    FormulaRewriteError, compute_order, compute_row_map, rewrite_formula)


def check(cond, label):
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def main() -> int:
    checks = []
    a = AxisAction(frame=200, kind="技能", excel_row=2)
    b = AxisAction(frame=100, frame2=150, kind="技能", excel_row=3)
    c = AxisAction(frame=150, kind="技能", excel_row=4)
    d = AxisAction(frame=300, kind="技能", excel_row=5)
    checks.append(check(compute_order([a, b, c, d]) == [3, 4, 2, 5],
                        f"分组排序 = {compute_order([a, b, c, d])}"))
    order, row_map = compute_row_map([a, b, c, d], 6)
    checks.append(check(order == [3, 4, 2, 5]
                        and row_map[3] == 2 and row_map[4] == 3
                        and row_map[2] == 4 and row_map[5] == 5
                        and row_map[6] == 6,
                        f"行映射 = {row_map}"))

    # 无操作 / 暂停 行也要按帧号一起排，不能被丢到表尾。
    mixed = [
        AxisAction(frame=1020, kind="技能", excel_row=2),
        AxisAction(frame=600, kind="技能", excel_row=3),
        AxisAction(frame=930, kind="无操作", excel_row=6),
        AxisAction(frame=950, kind="暂停", excel_row=7),
    ]
    order = compute_order(mixed)
    checks.append(check(order == [3, 6, 7, 2],
                        f"无操作/暂停一起排 = {order}"))
    # 没有帧号的行不参与排序（留在表尾），否则会顶到表头。
    order0 = compute_order(mixed + [AxisAction(frame=0, kind="无操作",
                                              excel_row=8)])
    checks.append(check(8 not in order0, f"0 帧行不进排序 = {order0}"))

    f = rewrite_formula("=NOT(K3)", "作战记录", "作战记录", {3: 8})
    checks.append(check(f == "=NOT(K8)", f"同行引用 = {f}"))
    f = rewrite_formula("=AND('作战记录'!K3<>\"\",'作战记录'!$K$3=TRUE)",
                        "其他", "作战记录", {3: 8})
    checks.append(check(f == "=AND('作战记录'!K8<>\"\",'作战记录'!$K$8=TRUE)",
                        f"跨表引用 = {f}"))
    try:
        rewrite_formula("=SUM(K3:K5)", "作战记录", "作战记录",
                        {3: 8, 4: 9, 5: 2})
        ok = False
    except FormulaRewriteError:
        ok = True
    checks.append(check(ok, "不连续区域引用会报错"))

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
