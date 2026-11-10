"""三行半透明帧窗口（独立进程，由 axis_console 启动）。

读取 out/showframe_state.json，显示当前帧号 / 当前时间 / 费用条。
右键或 Esc 关闭；控制台退出（含直接关窗口）时自动跟着关。
"""

from __future__ import annotations

import ctypes
import json
import os
import sys
import tkinter as tk
from ctypes import wintypes

STATE_PATH = (sys.argv[1] if len(sys.argv) > 1
              else os.path.join("out", "showframe_state.json"))
PARENT_PID = int(sys.argv[2]) if len(sys.argv) > 2 else 0
REFRESH_MS = 10
# 每 20 次刷新（约 0.2s）看一眼父进程还在不在。
PARENT_CHECK_EVERY = 20

_k32 = ctypes.WinDLL("kernel32", use_last_error=True)
_k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
_k32.OpenProcess.restype = wintypes.HANDLE
_k32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
_k32.WaitForSingleObject.restype = wintypes.DWORD
_k32.CloseHandle.argtypes = [wintypes.HANDLE]
_k32.CloseHandle.restype = wintypes.BOOL
_SYNCHRONIZE = 0x00100000
_WAIT_TIMEOUT = 0x102


def parent_alive() -> bool:
    """父进程（axis_console）还在不在。

    控制台被直接关掉时不会走 quit()，子进程不能靠 terminate 收尾，
    只能自己盯着父进程；父进程一没，这个窗口跟着关。
    """
    if PARENT_PID <= 0:
        return True
    handle = _k32.OpenProcess(_SYNCHRONIZE, False, PARENT_PID)
    if not handle:
        return False
    try:
        return _k32.WaitForSingleObject(handle, 0) == _WAIT_TIMEOUT
    finally:
        _k32.CloseHandle(handle)


def read_state():
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        alpha = float(data.get("alpha", 0.55))
        alpha = max(0.10, min(1.0, alpha))
        return (str(data.get("frame", "—")),
                str(data.get("time", "—")),
                str(data.get("cost", "—")),
                alpha)
    except Exception:
        return "—", "—", "—", None


def main() -> int:
    root = tk.Tk()
    root.overrideredirect(True)
    root.attributes("-topmost", True)
    root.attributes("-alpha", 0.55)
    root.configure(bg="#101010")
    root.geometry("+40+40")

    labels = []
    rows = []
    for text in ("当前帧号", "当前时间", "费用条"):
        row = tk.Frame(root, bg="#101010")
        row.pack(fill="x", padx=10, pady=1)
        tk.Label(row, text=text, fg="#B8B8B8", bg="#101010",
                 font=("Microsoft YaHei", 10)).pack(side="left")
        val = tk.Label(row, text="—", fg="#FFFFFF", bg="#101010",
                       font=("Consolas", 14, "bold"))
        val.pack(side="right")
        labels.append(val)
        rows.append(row)

    drag = {"x": 0, "y": 0}

    def press(event):
        drag["x"], drag["y"] = event.x_root, event.y_root

    def move(event):
        dx = event.x_root - drag["x"]
        dy = event.y_root - drag["y"]
        drag["x"], drag["y"] = event.x_root, event.y_root
        root.geometry(f"+{root.winfo_x() + dx}+{root.winfo_y() + dy}")

    def close(_event=None):
        root.destroy()

    for w in [root] + rows + labels:
        w.bind("<Button-1>", press)
        w.bind("<B1-Motion>", move)
        w.bind("<Button-3>", close)
    root.bind("<Escape>", close)

    applied_alpha = {"value": None}
    shown = {"frame": None, "time": None, "cost": None}
    ticks = {"n": 0}

    def refresh():
        try:
            ticks["n"] += 1
            if ticks["n"] % PARENT_CHECK_EVERY == 0 and not parent_alive():
                root.destroy()
                return
            frame, time_text, cost, alpha = read_state()
            if frame != shown["frame"]:
                labels[0].config(text=frame)
                shown["frame"] = frame
            if time_text != shown["time"]:
                labels[1].config(text=time_text)
                shown["time"] = time_text
            if cost != shown["cost"]:
                labels[2].config(text=cost)
                shown["cost"] = cost
            if alpha is not None and alpha != applied_alpha["value"]:
                root.attributes("-alpha", alpha)
                applied_alpha["value"] = alpha
            root.after(REFRESH_MS, refresh)
        except tk.TclError:
            return

    refresh()
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
