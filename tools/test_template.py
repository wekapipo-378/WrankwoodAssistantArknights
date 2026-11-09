import os
import sys

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from waa.core.axis import read_axis  # noqa: E402

src = os.path.join(ROOT, "axes", "WAA_操作轴模板.xlsx")
dst = os.path.join(ROOT, "out", "test_template.xlsx")
wb = load_workbook(src)
ex = wb["示例-1-7"]
ws = wb["作战记录"]
for r in range(2, 7):
    for c in range(1, 8):
        ws.cell(row=r, column=c, value=ex.cell(row=r, column=c).value)
wb.save(dst)
wb.close()

axis = read_axis(dst)
print("actions", len(axis.actions), "map", axis.map_code,
      "warnings", axis.warnings)
for a in axis.actions:
    print(a, "tile", a.tile_pos, "side", a.view_side, "front", a.view_front)
