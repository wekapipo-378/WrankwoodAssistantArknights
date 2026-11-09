"""生成用于手工验收的测试轴。

从 waa/templates/WAA_标准模板.xlsx 复制，写入 axes/测试用例。
运行测试后 Excel 会被原地写结果；重新运行本脚本即可恢复干净副本。
"""

from __future__ import annotations

import os
import shutil

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(ROOT, "waa", "templates", "WAA_标准模板.xlsx")
OUT_DIR = os.path.join(ROOT, "axes", "测试用例")


def set_setting(ws, label: str, value) -> None:
    """设置表在 M 列写键、N 列写值。"""
    for row in range(1, ws.max_row + 1):
        if str(ws.cell(row, 13).value or "").strip() == label:
            ws.cell(row, 14).value = value
            return
    raise KeyError(f"模板里没有设置项: {label}")


def clear_actions(ws) -> None:
    for row in range(2, ws.max_row + 1):
        for col in range(1, 12):
            ws.cell(row, col).value = None


def write_actions(ws, actions: list[dict]) -> None:
    clear_actions(ws)
    for i, item in enumerate(actions, start=2):
        ws.cell(i, 2).value = item.get("frame")
        ws.cell(i, 3).value = item.get("frame2")
        ws.cell(i, 4).value = item.get("kind")
        ws.cell(i, 5).value = item.get("oper")
        ws.cell(i, 6).value = item.get("pos")
        ws.cell(i, 7).value = item.get("direction")
        ws.cell(i, 8).value = item.get("alias")
        ws.cell(i, 9).value = item.get("note")
        if "enabled" in item:
            ws.cell(i, 10).value = item["enabled"]


def base_settings(ws, map_code: str = "1-7",
                  rates=(1,), default_2x=False,
                  start_pause=False, end_pause=False) -> None:
    set_setting(ws, "关卡代号", map_code)
    set_setting(ws, "回费速度", rates[0])
    # 额外回费候选放在 N/O/P；模板本身可能没有这些旧值，先清空再写。
    for col in range(15, 19):
        ws.cell(3, col).value = None
    for col, value in enumerate(rates, start=14):
        ws.cell(3, col).value = value
    set_setting(ws, "默认2倍速", default_2x)
    set_setting(ws, "开始前暂停", start_pause)
    set_setting(ws, "结束后暂停", end_pause)
    set_setting(ws, "打轴时暂停", True)
    set_setting(ws, "1x操作提前量", 4)
    set_setting(ws, "2x操作提前量", 5)
    set_setting(ws, "低回费盲步进", True)


def make(name: str, actions: list[dict], map_code: str = "1-7",
         rates=(1,), default_2x=False, start_pause=False,
         end_pause=False) -> str:
    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, name)
    shutil.copy2(TEMPLATE, path)
    wb = load_workbook(path)
    ws = wb["作战记录"]
    base_settings(ws, map_code, rates, default_2x, start_pause, end_pause)
    write_actions(ws, actions)
    wb.save(path)
    wb.close()
    print(path)
    return path


def main() -> None:
    make("01_全功能_1-7.xlsx", [
        {"frame": 0, "kind": "回费速度", "oper": "自动",
         "note": "自动回费"},
        {"frame": 30, "kind": "部署", "oper": "斑点", "pos": "D2",
         "direction": "右", "note": "单帧部署"},
        {"frame": 60, "kind": "部署", "oper": "Castle-3", "pos": "D5",
         "direction": "下", "alias": "近卫车"},
        {"frame": 90, "kind": "技能", "oper": "斑点",
         "note": "无坐标，用之前部署位置"},
        {"frame": 120, "frame2": 150, "kind": "技能", "oper": "Castle-3",
         "note": "区间技能"},
        {"frame": 180, "kind": "无操作", "note": "空操作"},
        {"frame": 210, "kind": "暂停", "note": "暂停行"},
        {"frame": 240, "kind": "撤退", "oper": "斑点",
         "note": "撤退"},
        {"frame": 270, "kind": "部署", "oper": "芬", "pos": "D3",
         "direction": "左", "enabled": False,
         "note": "禁用行，不应执行"},
        {"frame": 300, "kind": "部署", "oper": "芬", "pos": "E3",
         "direction": "右", "note": "同帧第一条"},
        {"frame": 300, "kind": "部署", "oper": "克洛丝", "pos": "F3",
         "direction": "下", "note": "同帧第二条"},
        {"frame": 330, "kind": "回费速度", "oper": "50%",
         "note": "手动回费"},
        {"frame": 360, "kind": "回费速度", "oper": "自动",
         "note": "恢复自动"},
    ], rates=(1, 0.5, 1.1))

    make("02_回费与负费.xlsx", [
        {"frame": 0, "kind": "回费速度", "oper": "100%"},
        {"frame": 60, "kind": "回费速度", "oper": "50%"},
        {"frame": 120, "kind": "回费速度", "oper": "自动"},
        {"frame": 180, "kind": "回费速度", "oper": "110%"},
        {"frame": 240, "frame2": 300, "kind": "回费速度", "oper": "100%",
         "note": "区间回费"},
        {"frame": 360, "kind": "回费速度", "oper": "自动"},
    ], rates=(1, 0.5, 1.1))

    make("03_边界错误.xlsx", [
        {"frame": 0, "kind": "回费速度", "oper": "自动"},
        {"frame": 30, "kind": "部署", "oper": "斑点",
         "note": "缺坐标和朝向"},
        {"frame": 60, "kind": "部署", "oper": "Castle-3", "pos": "D5",
         "note": "缺朝向"},
        {"frame": 90, "kind": "技能", "oper": "不存在干员",
         "note": "找不到位置"},
        {"frame": 120, "kind": "飞升", "oper": "斑点", "pos": "D2",
         "direction": "右", "note": "未知操作类型"},
        {"frame": 150, "frame2": 100, "kind": "无操作",
         "note": "帧号2小于帧号1"},
        {"kind": "部署", "oper": "芬", "pos": "E3",
         "direction": "右", "note": "缺帧号1"},
        {"kind": "无操作", "note": "无帧号空操作"},
        {"frame": 300, "kind": "部署", "oper": "芬", "pos": "E3",
         "direction": "右", "enabled": False, "note": "禁用"},
    ], rates=(1,))

    make("04_排序与同帧.xlsx", [
        {"frame": 300, "kind": "部署", "oper": "斑点", "pos": "D2",
         "direction": "右", "note": "乱序第一条"},
        {"frame": 100, "kind": "部署", "oper": "Castle-3", "pos": "D5",
         "direction": "下", "alias": "近卫车", "note": "乱序第二条"},
        {"frame": 200, "kind": "技能", "oper": "斑点", "note": "乱序技能"},
        {"frame": 200, "kind": "部署", "oper": "芬", "pos": "E3",
         "direction": "右", "note": "同帧 A"},
        {"frame": 200, "kind": "部署", "oper": "克洛丝", "pos": "F3",
         "direction": "下", "note": "同帧 B"},
        {"frame": 0, "kind": "回费速度", "oper": "自动"},
        {"frame": 400, "kind": "无操作", "note": "空操作"},
    ], rates=(1,))

    make("05_F5写入.xlsx", [
        {"frame": 0, "kind": "回费速度", "oper": "自动"},
        {"frame": 30, "kind": "部署", "oper": "斑点", "pos": "D2",
         "direction": "右", "note": "F5 写入目标"},
    ], rates=(1,))

    make("06_中途与偏移.xlsx", [
        {"frame": 0, "kind": "回费速度", "oper": "自动"},
        {"frame": 30, "kind": "部署", "oper": "斑点", "pos": "D2",
         "direction": "右"},
        {"frame": 60, "kind": "部署", "oper": "Castle-3", "pos": "D5",
         "direction": "下", "alias": "近卫车"},
        {"frame": 90, "kind": "技能", "oper": "斑点"},
        {"frame": 120, "kind": "撤退", "oper": "斑点"},
    ], rates=(1,))

    make("07_无地图.xlsx", [
        {"frame": 0, "kind": "回费速度", "oper": "自动"},
        {"frame": 30, "kind": "部署", "oper": "斑点", "pos": "D2",
         "direction": "右", "note": "没有地图数据"},
    ], map_code="", rates=(1,))


if __name__ == "__main__":
    main()
