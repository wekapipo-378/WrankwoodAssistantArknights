"""作战记录物理排序 + 公式引用重写。

排序：按帧号1 升序分组；重叠区间并组；组内按 Excel 行号；组间按最早帧号。
公式：只改引用“作战记录”被移动行的单元格；用 openpyxl Tokenizer，
遇到无法安全改写的公式就抛错，绝不带着错公式继续。
"""

from __future__ import annotations

import re

from openpyxl import load_workbook
from openpyxl.formula import Tokenizer

TARGET_SHEET = "作战记录"
SORT_COLUMNS = 11          # A..K 是操作表；L 用作临时排序列


class FormulaRewriteError(Exception):
    pass


def _interval(action) -> tuple[int, int]:
    end = action.frame2 if action.frame2 is not None else action.frame
    return int(action.frame), int(end)


def compute_order(actions) -> list[int]:
    """返回操作行的排序顺序（Excel 行号）。

    **不按操作类型过滤**：部署/技能/撤退/暂停/无操作/回费速度，以及以后新加
    的类型，一律按帧号一起排（用户要求“全排序，不用看操作”）。
    唯一例外是没有帧号（frame<=0）的行——它们交给 `compute_row_map` 按原
    顺序排到表尾，否则会顶到表头，把真正的第一条操作挤下去。
    """
    rows = [a for a in actions
            if getattr(a, "excel_row", 0)
            and int(getattr(a, "frame", 0) or 0) > 0]
    rows.sort(key=lambda a: (_interval(a)[0], a.excel_row))
    groups: list[list] = []
    group_end = None
    for a in rows:
        start, end = _interval(a)
        if groups and start <= group_end:
            groups[-1].append(a)
            group_end = max(group_end, end)
        else:
            groups.append([a])
            group_end = end
    order: list[int] = []
    for group in groups:
        group.sort(key=lambda a: a.excel_row)
        order.extend(int(a.excel_row) for a in group)
    return order


def compute_row_map(actions, last_row: int):
    """返回 (order_rows, row_map)：旧行 -> 新行。空行排在操作行之后。"""
    order_rows = compute_order(actions)
    row_map = {row: 2 + i for i, row in enumerate(order_rows)}
    empty = [r for r in range(2, last_row + 1) if r not in row_map]
    base = 2 + len(order_rows)
    for j, row in enumerate(empty):
        row_map[row] = base + j
    return order_rows, row_map


_SHEET = r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_一-鿿]+))!"
_CELL = r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)"
_REF_RE = re.compile(rf"^(?:{_SHEET})?{_CELL}(?::{_CELL})?$")


def _split_ref(operand: str):
    m = _REF_RE.match(operand)
    if not m:
        return None
    sheet = m.group(1) or m.group(2)
    if sheet:
        sheet = sheet.replace("''", "'")
    start = (m.group(3), m.group(4), m.group(5), int(m.group(6)))
    end = None
    if m.group(7) is not None:
        end = (m.group(7), m.group(8), m.group(9), int(m.group(10)))
    return sheet, start, end


def _cell(col_abs: str, col: str, row_abs: str, row: int) -> str:
    return f"{col_abs}{col.upper()}{row_abs}{row}"


def rewrite_formula(formula: str, formula_sheet: str,
                    target_sheet: str, row_map: dict[int, int]) -> str:
    """重写公式里指向 target_sheet 被移动行的单元格引用。"""
    upper = formula.upper()
    if "INDIRECT(" in upper or "OFFSET(" in upper:
        raise FormulaRewriteError(
            f"不支持 INDIRECT/OFFSET：{formula_sheet}!{formula}")
    out = []
    try:
        tokens = Tokenizer(formula).items
    except Exception as exc:
        raise FormulaRewriteError(f"公式无法解析：{formula!r} ({exc})")
    for tok in tokens:
        value = tok.value
        if tok.type != "OPERAND" or not value:
            out.append(value)
            continue
        if "!" in value:
            parsed = _split_ref(value)
            if parsed is None:
                raise FormulaRewriteError(
                    f"无法识别的跨表引用：{formula_sheet}!{formula}")
        else:
            parsed = _split_ref(value)
        if parsed is None:
            out.append(value)
            continue
        sheet, start, end = parsed
        use_sheet = sheet or formula_sheet
        if use_sheet != target_sheet:
            out.append(value)
            continue
        prefix = value.split("!", 1)[0] + "!" if "!" in value else ""
        c1, col1, r1a, r1 = start
        nr1 = row_map.get(r1, r1)
        if end is None:
            out.append(prefix + _cell(c1, col1, r1a, nr1))
            continue
        c2, col2, r2a, r2 = end
        nr2 = row_map.get(r2, r2)
        if (nr1 - r1) != (nr2 - r2):
            raise FormulaRewriteError(
                f"区域引用排序后不连续，无法安全改写："
                f"{formula_sheet}!{formula}")
        out.append(f"{prefix}{_cell(c1, col1, r1a, nr1)}:"
                   f"{_cell(c2, col2, r2a, nr2)}")
    result = "".join(out)
    return "=" + result if formula.startswith("=") else result


def collect_formula_rewrites(path: str, row_map: dict[int, int],
                             last_row: int) -> list[dict]:
    """读整本工作簿公式，返回排序后要写回的公式位置/文本。"""
    wb = load_workbook(path, data_only=False)
    rewrites: list[dict] = []
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f" or not cell.value:
                        continue
                    formula = str(cell.value)
                    if not formula.startswith("="):
                        formula = "=" + formula
                    # 其他表且完全没提到作战记录的公式与本次排序无关。
                    if (ws.title != TARGET_SHEET
                            and TARGET_SHEET not in formula
                            and f"'{TARGET_SHEET}'" not in formula):
                        continue
                    new_formula = rewrite_formula(
                        formula, ws.title, TARGET_SHEET, row_map)
                    new_row = cell.row
                    if ws.title == TARGET_SHEET and cell.column <= SORT_COLUMNS:
                        new_row = row_map.get(cell.row, cell.row)
                    rewrites.append({
                        "sheet": ws.title,
                        "row": int(new_row),
                        "col": int(cell.column),
                        "formula": new_formula,
                    })
    finally:
        wb.close()
    return rewrites
