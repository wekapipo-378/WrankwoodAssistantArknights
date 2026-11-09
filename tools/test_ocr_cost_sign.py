"""离线测试：运行信息“当前费用”的负号解析。"""

from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa.core.ocr_name import _parse_hud  # noqa: E402


def check(cond, label):
    print(("PASS " if cond else "FAIL ") + label)
    return cond


def main() -> int:
    checks = []
    cases = [
        (["-10"], -10, "ASCII 负号"),
        (["−10"], -10, "数学负号"),
        (["－10"], -10, "全角负号"),
        (["费用：-10"], -10, "带“费用”前缀"),
        (["10"], 10, "正数不变"),
        (["剩余可放置角色 6", "-10"], -10, "跳过剩余数量行"),
        (["-", "10"], -10, "负号单独一行"),
    ]
    for lines, want, label in cases:
        got = _parse_hud([], lines)["cost"]
        checks.append(check(got == want, f"{label}: {lines} -> {got}"))

    # 剩余可放置角色仍要正常解析，不能被费用负号影响。
    hud = _parse_hud([], ["剩余可放置角色 6", "-10"])
    checks.append(check(hud["deploy_left"] == 6,
                        f"剩余可放置角色 = {hud['deploy_left']}"))

    print("RESULT:", "ALL PASS" if all(checks) else "HAS FAILURES")
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
