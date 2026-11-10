"""F5 打轴写入独立 Excel。

和主运行面板分开：F5 只往这个文件末尾追加“帧号1”，不改主轴的
动作表。文件被 Excel 占用时保存会失败，由调用方报错，不会偷偷
回退到主轴。
"""

from __future__ import annotations

import threading
from copy import copy

from openpyxl import load_workbook


class MarkWriter:
    def __init__(self, path: str):
        self.path = str(path)
        self._lock = threading.Lock()
        self._wb = load_workbook(self.path)
        if "作战记录" not in self._wb.sheetnames:
            self._wb.close()
            raise ValueError("write 文件缺少“作战记录”工作表")
        self._ws = self._wb["作战记录"]

    def add_axis_row(self, frame: int) -> int:
        with self._lock:
            last = 1
            for r in range(2, int(self._ws.max_row) + 1):
                b = self._ws.cell(r, 2).value
                d = self._ws.cell(r, 4).value
                if (b not in (None, "")) or (d not in (None, "")):
                    last = r
            row = last + 1
            # 只写值会让新行套用工作表默认样式（微软雅黑 10），和模板数据行
            # 的宋体 11 不一样。把样式从第一条记录（第 2 行）整行复制过来，
            # 不取 last——last 可能是之前 F5 加的行，风格会一路传下去。
            style_row = 2 if self._ws.cell(2, 2).value not in (None, "") else last
            if style_row >= 2:
                for col in range(1, 12):
                    src = self._ws.cell(style_row, col)
                    dst = self._ws.cell(row, col)
                    dst._style = copy(src._style)
            self._ws.cell(row, 2).value = int(frame)
            self._wb.save(self.path)
            return row

    def close(self) -> None:
        try:
            self._wb.close()
        except Exception:
            pass
