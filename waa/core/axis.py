"""读取 prts-plus 格式的 Excel 轴，并把坐标换算成屏幕比例。

坐标约定沿用 prts-plus：
    convert_position("D2", height, width) -> (row=1, col=3)
    取视图位置时是 view[col][row]（注意不是 view[row][col]）。
"""

from __future__ import annotations

import copy
import json
import math
import os
from dataclasses import dataclass, field
from typing import Any

import numpy as np
from openpyxl import load_workbook

from .resources import LEVELS_PATH, RESOURCE_LOCK

COLUMNS = {
    "费用": "cost",
    "帧号": "tick",
    "帧号1": "tick1",
    "帧号2": "tick2",
    "操作": "kind",
    "单位": "oper",
    "干员": "oper",   # 旧模板列名，读取时兼容
    "坐标": "pos",
    "朝向": "direction",
    "简称": "alias",
    "备注": "note",
    "是否运行": "enabled",
    "关卡名": "map",
}


@dataclass
class AxisAction:
    frame: int
    kind: str
    oper: str | None = None
    pos: str | None = None
    direction: str | None = None
    alias: str | None = None
    note: str | None = None
    tile_pos: tuple[int, int] | None = None
    view_side: tuple[float, float] | None = None
    view_front: tuple[float, float] | None = None
    frame2: int | None = None   # 区间末帧；None = 只尝试 frame 这一帧
    enabled: bool = True        # “是否运行”列；False 时整行跳过
    excel_row: int = 0   # 这一行在“作战记录”里的 Excel 行号（1 基），给状态回写用
    # “操作=回费速度”专用：rate_auto=True 表示恢复自动检测；
    # rate_percent 是手动回费速度（80 表示 80%）。
    rate_auto: bool = False
    rate_percent: float | None = None

    @property
    def end_frame(self) -> int:
        return self.frame2 if self.frame2 is not None else self.frame

    @property
    def frame_text(self) -> str:
        if self.frame2 is not None and self.frame2 != self.frame:
            return f"{self.frame}~{self.frame2}"
        return str(self.frame)

    def __str__(self) -> str:
        return (
            f"{self.frame_text} {self.kind} {self.oper or ''} "
            f"{self.pos or ''} {self.direction or ''} {self.alias or ''}"
        )


@dataclass
class Axis:
    path: str
    settings: dict[str, Any]
    actions: list[AxisAction]
    map_code: str | None = None
    map_data: dict | None = None
    warnings: list[str] = field(default_factory=list)
    columns: dict[str, int] = field(default_factory=dict)  # 表头 -> 0 基列号


_LEVEL_INDEX: dict[str, dict] | None = None
_LEVEL_NAME_INDEX: dict[str, dict] | None = None


def _load_level_index() -> tuple[dict[str, dict], dict[str, dict]]:
    """从本地 levels.json 建 code/name -> 地图对象索引，进程内缓存。"""
    global _LEVEL_INDEX, _LEVEL_NAME_INDEX
    with RESOURCE_LOCK:
        if _LEVEL_INDEX is not None and _LEVEL_NAME_INDEX is not None:
            return _LEVEL_INDEX, _LEVEL_NAME_INDEX
        if not os.path.exists(LEVELS_PATH):
            raise FileNotFoundError(
                f"找不到本地地图数据: {LEVELS_PATH}；"
                "请检查 out/ArknightsGameResource")
        with open(LEVELS_PATH, encoding="utf-8") as fh:
            levels = json.load(fh)
        index: dict[str, dict] = {}
        names: dict[str, dict] = {}
        for level in levels:
            if not isinstance(level, dict):
                continue
            code = str(level.get("code") or "").strip()
            name = str(level.get("name") or "").strip()
            if code:
                index.setdefault(code, level)
            if name:
                names.setdefault(name, level)
        _LEVEL_INDEX = index
        _LEVEL_NAME_INDEX = names
        return index, names


def clear_map_cache() -> None:
    global _LEVEL_INDEX, _LEVEL_NAME_INDEX
    with RESOURCE_LOCK:
        _LEVEL_INDEX = None
        _LEVEL_NAME_INDEX = None


def load_map_by_code(map_code: str) -> dict:
    with RESOURCE_LOCK:
        index, names = _load_level_index()
        key = str(map_code).strip()
        level = index.get(key) or names.get(key)
        if level is None:
            raise KeyError(f"本地 levels.json 里没有地图代号: {map_code}")
        return copy.deepcopy(level)


def convert_position(pos: str, height: int, width: int) -> tuple[int, int]:
    """D2 -> (1, 3)。与 prts-plus 保持一致（用 height 换算字母）。"""
    return (int(pos[1:]) - 1, height - 1 - (ord(pos[0]) - ord("A")))


def transform_map_to_view(level: dict, side: bool) -> list[list[tuple[float, float]]]:
    degree = math.pi / 180
    height, width = level["height"], level["width"]
    x, y, z = level["view"][1 if side else 0]

    transform = np.array([
        [1, 0, 0, -x],
        [0, 1, 0, -y],
        [0, 0, 1, -z],
        [0, 0, 0, 1],
    ])
    perspective = np.array([
        [9 / 16 / math.tan(20 * degree), 0, 0, 0],
        [0, 1 / math.tan(20 * degree), 0, 0],
        [0, 0, -(1000 + 0.3) / (1000 - 0.3), -(1000 * 0.3 * 2) / (1000 - 0.3)],
        [0, 0, -1, 0],
    ])
    rotate_x = np.array([
        [1, 0, 0, 0],
        [0, math.cos(30 * degree), -math.sin(30 * degree), 0],
        [0, -math.sin(30 * degree), -math.cos(30 * degree), 0],
        [0, 0, 0, 1],
    ])
    rotate_y = np.array([
        [math.cos(10 * degree), 0, math.sin(10 * degree), 0],
        [0, 1, 0, 0],
        [-math.sin(10 * degree), 0, math.cos(10 * degree), 0],
        [0, 0, 0, 1],
    ])
    final = perspective @ rotate_x @ rotate_y @ transform if side else perspective @ rotate_x @ transform

    out: list[list[tuple[float, float]]] = []
    for i in range(height):
        row: list[tuple[float, float]] = []
        for j in range(width):
            tile = level["tiles"][i][j]
            point = np.array([
                j - (width - 1) / 2.0,
                (height - 1) / 2.0 - i,
                tile["heightType"] * -0.4,
                1,
            ])
            view = final @ point
            view = view / view[3]
            view = (view + 1) / 2
            row.append((float(view[0]), float(1 - view[1])))
        out.append(row)
    return out


def _view_at(view: list[list[tuple[float, float]]], tile_pos: tuple[int, int]) -> tuple[float, float]:
    row, col = tile_pos
    # prts-plus 的取法：view[col][row]
    return view[col][row]


def _read_settings(ws) -> dict[str, Any]:
    """读取“设置”区，兼容两种模板：

    - WAA_full_1-7：设置键在 M 列、值在 N 列；
    - WAA_sample_1-7：设置键在 L 列、值在 M 列（“可出现回费速度”还会往右写多个值，
      例如 M3=1、N3=0.5，表示先 100% 后 50%）。

    找到第 1 行“设置”表头所在列，键读该列、值读它右边连续非空单元格。
    """
    settings_col = None
    for cell in ws[1]:
        if cell.value is not None and str(cell.value).strip() == "设置":
            settings_col = cell.column
            break
    if settings_col is None:
        settings_col = 13  # 旧模板默认 M

    out: dict[str, Any] = {}
    for r in range(1, ws.max_row + 1):
        key = ws.cell(row=r, column=settings_col).value
        if key is None or str(key).strip() == "":
            continue
        key = str(key).strip()
        values = []
        for c in range(settings_col + 1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is None or str(v).strip() == "":
                break
            values.append(v)
        if not values:
            out[key] = None
        elif key in ("可出现回费速度", "回费速度", "可能出现回费速度"):
            out[key] = values
        else:
            out[key] = values[0] if len(values) == 1 else values
    return out


def parse_rate_value(value):
    """解析“回费速度”行的单位：

    - “自动” / auto -> 字符串 "auto"
    - 80% / 0.8 / 80   -> 80.0
    - 其它/空白        -> None
    """
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    if s.lower() in ("自动", "auto"):
        return "auto"
    try:
        if s.endswith("%"):
            rate = float(s[:-1])
        else:
            f = float(s)
            rate = f * 100.0 if f <= 2.0 else f
    except ValueError:
        return None
    if 0.0 < rate <= 1000.0:
        return rate
    return None


def read_axis(path: str) -> Axis:
    wb = load_workbook(path, data_only=True, keep_vba=path.lower().endswith(".xlsm"))
    ws = wb["作战记录"]
    headers: dict[str, int] = {}
    for cell in ws[1]:
        if cell.value is not None:
            headers[str(cell.value).strip()] = cell.column - 1
    # 旧模板把这一列叫“干员”，新模板叫“单位”；统一按“单位”取。
    if "单位" not in headers and "干员" in headers:
        headers["单位"] = headers["干员"]

    settings = _read_settings(ws)

    raw_rows: list[dict[str, Any]] = []
    for row_idx, row in enumerate(
            ws.iter_rows(min_row=2, values_only=True), start=2):
        item = {}
        for name, idx in headers.items():
            if idx < len(row):
                item[name] = row[idx]
        # 空操作行也要保留：运行结果要写“无操作”。
        if all(item.get(k) in (None, "") for k in
               ("操作", "帧号1", "帧号2", "帧号", "备注")):
            continue
        item["_excel_row"] = row_idx
        raw_rows.append(item)
    # 关卡名可能只填在第一行；即使那一行没有“操作”也要读到。
    map_code = None
    map_col = headers.get("关卡名")
    if map_col is not None:
        for row in ws.iter_rows(min_row=2, values_only=True):
            if map_col < len(row) and row[map_col] not in (None, ""):
                map_code = str(row[map_col]).strip()
                break
    wb.close()

    # 兼容旧模板右侧设置里的“关卡代号”。
    if not map_code:
        map_code = settings.get("关卡代号") or settings.get("关卡名")
    map_data = load_map_by_code(str(map_code)) if map_code else None
    view_side = transform_map_to_view(map_data, True) if map_data else None
    view_front = transform_map_to_view(map_data, False) if map_data else None
    height = map_data["height"] if map_data else 0
    width = map_data["width"] if map_data else 0

    def _text(raw: dict[str, Any], key: str) -> str | None:
        v = raw.get(key)
        if v is None or str(v).strip() == "":
            return None
        return str(v).strip()

    def _int(raw: dict[str, Any], key: str) -> int | None:
        v = raw.get(key)
        if v is None or str(v).strip() == "":
            return None
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return None

    def _bool(raw: dict[str, Any], key: str, default: bool = True) -> bool:
        v = raw.get(key)
        if v is None or str(v).strip() == "":
            return default
        if isinstance(v, bool):
            return v
        s = str(v).strip().lower()
        if s in ("true", "1", "yes", "y", "是", "开"):
            return True
        if s in ("false", "0", "no", "n", "否", "关"):
            return False
        return default

    # 先收集所有“简称 -> 干员”，这样乱序行也能正确解析。
    alias_map: dict[str, str] = {}
    for raw in raw_rows:
        alias = _text(raw, "简称")
        oper = _text(raw, "单位")
        if alias and oper:
            alias_map[alias] = oper

    actions: list[AxisAction] = []
    warnings: list[str] = []
    # 每个干员/简称的部署历史：(部署帧, 格子)；乱序时按帧找“最近一次部署”。
    deploy_history: dict[str, list[tuple[int, tuple[int, int]]]] = {}
    for raw in raw_rows:
        if (_text(raw, "操作") or "") != "部署":
            continue
        pos = _text(raw, "坐标")
        oper = _text(raw, "单位")
        frame = _int(raw, "帧号1")
        if frame is None:
            frame = _int(raw, "帧号")
        if not pos or not oper or frame is None:
            continue
        if oper in alias_map:
            oper = alias_map[oper]
        tile = convert_position(pos, height, width)
        keys = [oper]
        alias = _text(raw, "简称")
        if alias:
            keys.append(alias)
        for key in keys:
            deploy_history.setdefault(key, []).append((frame, tile))

    for raw in raw_rows:
        kind = _text(raw, "操作") or "无操作"
        frame1 = _int(raw, "帧号1")
        if frame1 is None:
            frame1 = _int(raw, "帧号")
        frame2 = _int(raw, "帧号2")
        oper = _text(raw, "单位")
        if oper and oper in alias_map:
            oper = alias_map[oper]
        action = AxisAction(
            frame=frame1 if frame1 is not None else 0,
            kind=kind,
            oper=oper,
            pos=_text(raw, "坐标"),
            direction=_text(raw, "朝向"),
            alias=_text(raw, "简称"),
            note=_text(raw, "备注"),
            excel_row=int(raw.get("_excel_row") or 0),
            enabled=_bool(raw, "是否运行", True),
        )
        if frame1 is None:
            if kind != "无操作":
                warnings.append(f"缺少帧号1: {action}")
                continue
        elif frame2 is not None:
            if frame2 < frame1:
                warnings.append(
                    f"帧号2 < 帧号1，按单帧处理: {action}")
            else:
                action.frame2 = frame2

        if kind in ("无操作", "暂停"):
            # 无操作只写结果；暂停行只需要帧号。
            actions.append(action)
            continue

        if kind == "回费速度":
            if frame2 is not None and frame2 < action.frame:
                warnings.append(f"回费速度帧号2 < 帧号1，已跳过: {action}")
                continue
            parsed = parse_rate_value(action.oper)
            if parsed is None:
                warnings.append(f"回费速度行缺少或无法识别单位: {action}")
                continue
            if parsed == "auto":
                action.rate_auto = True
            else:
                action.rate_percent = float(parsed)
            actions.append(action)
            continue

        if kind == "部署":
            if not action.oper or not action.pos or not action.direction:
                warnings.append(f"部署行缺字段: {action}")
                continue
            action.tile_pos = convert_position(action.pos, height, width)
            actions.append(action)
            continue

        # 技能/撤退：位置可以显式写，也可以找最近一次部署。
        if action.pos:
            action.tile_pos = convert_position(action.pos, height, width)
        else:
            candidates: list[tuple[int, tuple[int, int]]] = []
            for key in (action.oper, action.alias):
                if key:
                    candidates.extend(deploy_history.get(key, []))
            if candidates:
                prior = [x for x in candidates if x[0] <= action.frame]
                pick = (max(prior, key=lambda x: x[0]) if prior
                        else min(candidates, key=lambda x: x[0]))
                action.tile_pos = pick[1]
            else:
                warnings.append(f"技能/撤退找不到位置: {action}")
                continue
        actions.append(action)

    # 部署历史可能在技能行之后才出现，所以位置解析完再统一算视图坐标。
    for action in actions:
        if action.tile_pos is None or view_side is None:
            continue
        action.view_side = _view_at(view_side, action.tile_pos)
        action.view_front = _view_at(view_front, action.tile_pos)

    return Axis(
        path=path,
        settings=settings,
        actions=actions,
        map_code=str(map_code) if map_code else None,
        map_data=map_data,
        warnings=warnings,
        columns=headers,
    )
