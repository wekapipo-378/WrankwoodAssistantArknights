"""运行状态实时写回 Excel（COM）。

脚本以管理员权限运行，自己启动一个 Excel 实例。

普通模式打开**模板副本**：

    out/run_reports/<轴名>_<时间>_运行.xlsx

原模板文件不动。

原地模式（in_place=True）直接打开 axis.path 写回，不自动备份；
运行前先保存该文件。文件可以由本程序打开；如果用户已经用 Excel
打开且已保存，则优先连接现有 Excel，不再另开实例。写的内容：

    作战记录    A 当前执行箭头；J 每个动作的运行结果
    运行信息    B1 当前状态、B2 当前操作、B3 报错信息、B4 帧号、B5 回费速度、
                B6 倍速、B7 暂停状态、B8 当前选中；
                下方是 1-7 的格子示意图（地形、已部署干员、朝向、当前目标）

更新频率约 2~5 次/秒；F6 手动暂停后可以切到 Excel 看，再按 F6 回游戏。
普通模式下 Excel 打不开不影响打轴，只在控制台提示；原地模式下面板打不开
会直接停止（run_axis 检查 start() 的返回值）。结束时另写 out/run_status.json。
"""

from __future__ import annotations

import json
import os
import shutil
import threading
import time
import traceback
from datetime import datetime

from .ocr_name import HudOCR

# 注意：**不要在这里 import win32com/pythoncom**。
# 主线程还要用 windows_capture 的 WinRT 抓屏；pywin32 的 COM 初始化会和
# WinRT 抢 COM apartment，导致 grab_client 报 “Failed to initialize WinRT”。
# COM 只在 Excel 上报线程里延迟导入 + CoInitialize。

# 版式：
#   第 1 页 作战记录（和 WAA_sample_1-7 一样的列/设置）
#   第 2 页 地图（上面状态，下面地图）
REPORT_SHEETS = ("作战记录", "运行信息")
MAP_TITLE_ROW = 1
MAP_HEADER_ROW = 2
MAP_ROW0 = 3          # 地图格子起点（Excel 1 基）
MAP_COL0 = 2
STATUS_ROWS = (
    ("当前状态", "准备中"), ("当前操作", ""), ("报错信息", ""),
)
STATUS_FIELDS = (
    (5, "当前帧号"), (6, "当前时间"), (7, "费用条"),
    (8, "回费速度"), (9, "当前倍速"), (10, "当前费用"),
    (11, "剩余可放置角色"), (13, "当前选中"),
    (14, "攻击"), (15, "防御"), (16, "法抗"), (17, "阻挡"),
)


def _frame_time_text(frame: int) -> str:
    """全局帧 -> mm:ss:ff（30 逻辑帧/秒，秒 00~59）。"""
    frame = max(0, int(frame))
    total_sec = frame // 30
    mm, ss = divmod(total_sec, 60)
    ff = frame % 30
    return f"{mm:02d}:{ss:02d}:{ff:02d}"


def _cost_frame_text(last) -> str:
    """当前费用周期的 phase/cycle_len；第 11 个周期可能是 31 帧。"""
    if last is None:
        return ""
    phase = getattr(last, "phase", None)
    cycle_len = getattr(last, "cycle_len", 0)
    if phase is None or not cycle_len:
        return ""
    return f"{int(phase)}/{int(cycle_len)}"

TILE_COLORS = {
    "tile_forbidden": "595959",
    "tile_wall": "8B5A2B",
    "tile_floor": "D9D9D9",
    "tile_road": "F4E4C1",
    "tile_hole": "262626",
    "tile_start": "C6EFCE",
    "tile_end": "FFC7CE",
    "tile_flystart": "C6EFCE",
    "tile_telout": "DDEBF7",
    "tile_telin": "DDEBF7",
}
DEFAULT_TILE = "BFBFBF"
DEPLOYED_COLOR = "9DC3E6"
TARGET_COLOR = "FFD966"
ARROWS = {"右": "→", "左": "←", "上": "↑", "下": "↓"}


def _excel_color(hex_rgb: str) -> int:
    """RRGGBB -> Excel 的 BGR 整数。"""
    r = int(hex_rgb[0:2], 16)
    g = int(hex_rgb[2:4], 16)
    b = int(hex_rgb[4:6], 16)
    return r + (g << 8) + (b << 16)


def _to_bool(value, default: bool = False) -> bool:
    if value is None or str(value).strip() == "":
        return default
    if isinstance(value, bool):
        return value
    s = str(value).strip().lower()
    if s in ("true", "1", "yes", "y", "是", "开"):
        return True
    if s in ("false", "0", "no", "n", "否", "关"):
        return False
    # #N/A、#VALUE! 等公式错误或无法识别的文本一律不执行。
    return False


class ExcelReporter:
    def __init__(self, axis, rate_info: dict, report_dir: str,
                 enabled: bool = True, quit_on_stop: bool = False,
                 template_path: str | None = None,
                 in_place: bool = False, backup: bool = False):
        self.axis = axis
        self.rate_info = dict(rate_info or {})
        self.report_dir = report_dir
        self.enabled = bool(enabled)
        self.quit_on_stop = bool(quit_on_stop)
        # 普通模式固定用模板复制新文件；原地模式直接打开 axis.path。
        self.template_path = template_path or axis.path
        self.in_place = bool(in_place)
        self.backup = bool(backup)
        self.backup_path = ""
        self.start_error = ""
        self._win32_client = None
        self.report_path = ""
        self._json_path = os.path.join(
            os.path.dirname(report_dir), "run_status.json")

        self._lock = threading.Lock()
        self._compute_lock = threading.Lock()
        self._start_lock = threading.Lock()
        self.started = False
        self._stop = False
        self._ready = threading.Event()
        self._save_request = threading.Event()
        self._save_done = threading.Event()
        self._save_ok = False
        self._front_request = threading.Event()
        self._front_done = threading.Event()
        self._front_ok = False
        self._thread: threading.Thread | None = None
        self._last_com_warn_at = 0.0
        self._excel_ok = False
        self._attached = False
        self._app = None
        self._wb = None
        self._ws_rec = None
        self._ws_mid = None
        self._ws_map = None
        self._cols: dict[str, int] = {}

        self._status = "准备中"
        self._current_text = ""
        self._error = ""
        self._frame_text = ""
        self._time_text = ""
        self._cost_frame_text = ""
        self._rate_text = self._realtime_rate_text(False)
        self._speed_text = ""
        self._pause_text = ""
        self._selected_text = ""
        self._cost_text = ""
        self._deploy_left_text = ""
        self._attack_text = ""
        self._defense_text = ""
        self._resist_text = ""
        self._block_text = ""
        self._hud_read_at = 0.0
        self._finished = False
        self._ok = False

        self._current_index: int | None = None
        self._current_action = None
        self._results: dict[int, dict] = {}
        self._failures = 0
        self._order = 0
        self._enabled_read_at = 0.0
        self._result_dirty = False
        self._result_saved_at = 0.0
        # 同步屏障：主线程发请求，Excel 线程写结果/重算/读 J 后回传。
        self._sync_lock = threading.Lock()
        self._sync_gate = threading.Lock()
        self._sync_kind: str | None = None
        self._sync_index: int | None = None
        self._sync_ok: bool | None = None
        self._sync_payload = None
        self._sync_result = None
        self._sync_done = threading.Event()
        self._wake = threading.Event()
        self._deployed: dict[str, tuple] = {}
        self._dirty = True
        self._map_dirty = True
        self._arrow_row: int | None = None
        self._written: dict[str, str] = {}
        self._terrain: dict[tuple[int, int], str] = {}
        self._marked: set[tuple[int, int]] = set()
        self._map_signature = None
        self._map_title_row = MAP_TITLE_ROW
        self._map_header_row = MAP_HEADER_ROW
        self._map_row0 = MAP_ROW0

        self._counter = None
        self._executor = None
        self._ocr = HudOCR()
        self._started_at = time.time()

    # ---------- 对外接口 ----------
    def _realtime_rate_text(self, neg: bool = False) -> str:
        """地图页显示**当前**回费速度；不是设置页的“可能出现速度”列表。

        正常回费显示设置的基准速度（默认 100%）；局内欠费时显示
        基准速度的一半，并明确标“欠费”。
        手动回费速度生效时显示该值，并标“手动”。
        """
        runtime = getattr(self, "_counter", None)
        counter = getattr(runtime, "counter", None)
        # 正式 WAACounter 有 in_battle；不在关卡时不要显示上一局的欠费/手动状态。
        if runtime is not None and not bool(getattr(runtime, "in_battle", True)):
            base = float(self.rate_info.get("base", 100.0))
            return f"{base:g}%"
        manual = getattr(counter, "manual_rate", None)
        if manual is not None:
            current = float(manual) / 2.0 if neg else float(manual)
            suffix = "，欠费" if neg else ""
            return f"{current:g}%（手动{suffix}）"
        # 自动检测会改 counter.rate_percent；显示当前实际采用的档位。
        detected = getattr(counter, "rate_percent", None)
        base = (float(detected) if detected
                else float(self.rate_info.get("base", 100.0)))
        if neg:
            return f"{base / 2.0:g}%（欠费）"
        return f"{base:g}%"

    def start(self, wait_ready: bool = True) -> bool:
        """启动 Excel 面板；返回 False 表示面板没打开。

        普通模式：复制模板到 report_dir，原轴不动。
        原地模式：直接打开 axis.path 并实时写回（默认不备份）。

        正常开局由 run_axis 在编队界面、点“开始行动”之前调用；这时
        战斗还没开始，打开 Excel 不会丢帧。已经在局内时也允许由第一次
        F6 手动暂停兜底调用。
        """
        with self._start_lock:
            if self.started:
                return bool(self._excel_ok)
            old = self._thread
            if old is not None and old.is_alive():
                # 不能在上一条 Excel 面板线程还没退出时再开一条：
                # 两条线程会共享 _wb/_ws_*，COM 对象跨公寓使用就会
                # RPC_E_WRONG_THREAD。
                print("[excel] 上一次 Excel 面板线程还没退出，等待它结束…")
                self._stop = True
                old.join(timeout=2.0)
                if old.is_alive():
                    self.start_error = (
                        "上一次 Excel 面板线程还没退出；请稍后再试")
                    print(f"[excel] {self.start_error}")
                    return False
            self._thread = None
            self._stop = False
            self._ready.clear()
            self._excel_ok = False
            self.start_error = ""
            self._attached = False
            self._app = None
            self._wb = None
            self._ws_rec = None
            self._ws_mid = None
            self._ws_map = None
            self.started = True

        if not self.enabled:
            print("[excel] 面板未启用（--no-excel 或缺少 pywin32）")
            self.start_error = "面板未启用"
            self.started = False
            return False

        if self.in_place:
            self.report_path = os.path.abspath(self.axis.path)
            if not os.path.exists(self.report_path):
                self.start_error = f"找不到轴文件：{self.report_path}"
                print(f"[excel] {self.start_error}")
                self.enabled = False
                self.started = False
                return False
            # 文件已经被用户用 Excel 打开时，不再另开一个实例，而是
            # 尝试连接现有 Excel（见 _open）。只有连接失败才报错。
            if self.backup:
                try:
                    self.backup_path = self._backup_path()
                    shutil.copy2(self.report_path, self.backup_path)
                    print(f"[excel] 已备份原文件：{self.backup_path}")
                except OSError as e:
                    self.start_error = f"备份失败：{e}"
                    print(f"[excel] {self.start_error}；已停止")
                    self.enabled = False
                    self.started = False
                    return False
        else:
            os.makedirs(self.report_dir, exist_ok=True)
            stem = os.path.splitext(os.path.basename(self.axis.path))[0]
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.report_path = os.path.join(
                self.report_dir, f"{stem}_{stamp}_运行.xlsx")
            base = (self.template_path
                    if os.path.exists(self.template_path) else self.axis.path)
            try:
                shutil.copy2(base, self.report_path)
            except OSError as e:
                self.start_error = f"复制模板失败：{e}"
                print(f"[excel] {self.start_error}，面板关闭")
                self.enabled = False
                self.started = False
                return False

        try:
            os.makedirs(os.path.dirname(self._json_path), exist_ok=True)
        except OSError:
            pass
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        if wait_ready:
            self._ready.wait(20.0)
        if self.in_place:
            print(f"[excel] 运行面板（原文件）：{self.report_path}")
        else:
            print(f"[excel] 运行面板：{self.report_path}")
        if not self._excel_ok:
            print(f"[excel] 面板没打开：{self.start_error or 'Excel/COM 初始化失败'}")
            # 允许下次 F6 再试（比如用户刚刚才关掉占用文件的 Excel）。
            self.started = False
        return bool(self._excel_ok)

    def _backup_path(self) -> str:
        stem, ext = os.path.splitext(self.report_path)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = f"{stem}_备份_{stamp}{ext}"
        n = 1
        while os.path.exists(path):
            path = f"{stem}_备份_{stamp}_{n}{ext}"
            n += 1
        return path

    def bind_runtime(self, runtime, executor) -> None:
        self._counter = runtime
        self._executor = executor

    def save_workbook(self, timeout: float = 5.0) -> bool:
        """让 Excel 线程保存当前工作簿；F6 恢复前重读操作用。"""
        if not self.enabled or not self.started or not self._excel_ok:
            return False
        if self._wb is None:
            return False
        self._save_ok = False
        self._save_done.clear()
        self._save_request.set()
        if not self._save_done.wait(timeout):
            print("[excel] 保存工作簿超时")
            return False
        return bool(self._save_ok)

    def bring_to_front(self, timeout: float = 3.0) -> bool:
        """F6 暂停后把 Excel 窗口切到前台；COM 操作在 reporter 线程做。

        返回**真实结果**：窗口真的到前台才 True。旧代码 finally 里无条件 set
        `_front_done`，即使 SetForegroundWindow 失败也返回 True，导致调用方
        （_manual_hold）打印「Excel 已前置」的假成功（2026-11-09 08:23 轮）。
        """
        if not self.enabled or not self.started or not self._excel_ok:
            return False
        self._front_ok = False
        self._front_done.clear()
        self._front_request.set()
        self._front_done.wait(timeout)
        return self._front_ok

    def sync_read_enabled(self, index: int, timeout: float = 0.5):
        """同步读某行的“是否执行”，返回 True/False/None。"""
        return self._sync_call("read", index, None, timeout)

    def sync_commit_result(self, index: int, ok: bool | None,
                           timeout: float = 0.5) -> bool:
        """同步写某行的“执行结果”（True/False/空）并重算。"""
        return self._sync_call("commit", index, ok, timeout) is True

    def add_axis_row(self, frame: int, timeout: float = 2.0):
        """F5 打轴：在作战记录末尾加一行，帧号1 填当前帧。返回行号。"""
        return self._sync_call("add_row", 0, None, timeout, payload=int(frame))

    def sort_records(self, order_rows: list[int], last_row: int,
                     timeout: float = 15.0) -> bool:
        """在 Excel 线程里物理排序作战记录 A..K。"""
        return self._sync_call("sort", 0, None, timeout,
                               payload=(order_rows, last_row)) is True

    def apply_formula_rewrites(self, rewrites: list[dict],
                               timeout: float = 15.0) -> bool:
        """在 Excel 线程里把重写后的公式写回。"""
        return self._sync_call("formulas", 0, None, timeout,
                               payload=rewrites) is True

    def _sync_call(self, kind: str, index: int, ok: bool | None,
                   timeout: float, payload=None):
        # 主线程屏障和 F5 热键可能同时请求，串行化，避免互相覆盖请求字段。
        with self._sync_gate:
            return self._sync_call_inner(kind, index, ok, timeout, payload)

    def _sync_call_inner(self, kind: str, index: int, ok: bool | None,
                         timeout: float, payload=None):
        if not self.enabled or not self.started or not self._excel_ok:
            return None
        if self._wb is None:
            return None
        with self._sync_lock:
            self._sync_kind = kind
            self._sync_index = int(index)
            self._sync_ok = ok
            self._sync_payload = payload
            self._sync_result = None
            self._sync_done.clear()
        self._wake.set()
        if self._sync_done.wait(timeout):
            return self._sync_result
        print(f"[excel] 同步请求超时：{kind} 行 {index}")
        return None

    def set_error(self, msg: str) -> None:
        with self._lock:
            self._error = str(msg)
            self._dirty = True
            if self._current_index is not None:
                r = self._results.setdefault(self._current_index, {})
                action = self._current_action
                if action is not None:
                    r.setdefault("order", None)
                r["ok"] = False
                r["note"] = "失败"
                r["error"] = str(msg)
                self._map_dirty = True
                self._result_dirty = True

    def action_start(self, index: int, action) -> None:
        with self._lock:
            self._current_index = int(index)
            self._current_action = action
            self._current_text = f"[{int(index)}] {action}"
            self._dirty = True
            self._map_dirty = True

    def action_done(self, index: int, action, target,
                    actual: int | None, note: str = "") -> None:
        with self._lock:
            actual_i = int(actual) if actual is not None else None
            if note == "无操作":
                ok: bool | None = None
            elif note == "成功" or (note not in ("失败",) and actual_i is not None):
                ok = True
            else:
                ok = False
            order = None
            if ok is not None and getattr(action, "enabled", True):
                self._order += 1
                order = self._order
            self._results[int(index)] = {
                "ok": ok,
                "note": "无操作" if ok is None else ("成功" if ok else "失败"),
                "order": order,
                "action": action,
            }
            if ok is False:
                self._failures += 1
            self._result_dirty = True
            if actual_i is not None and action.kind == "部署" and action.oper:
                self._deployed[action.oper] = (
                    action.tile_pos, action.direction,
                    action.alias or action.oper)
            elif actual_i is not None and action.kind == "撤退" and action.oper:
                self._deployed.pop(action.oper, None)
            # 动作结束后不再把当前目标画成黄色，已部署的会显示蓝色
            self._current_action = None
            self._dirty = True
            self._map_dirty = True

    def record_simple_action(self, index: int, action, frame: int,
                             note: str = "成功") -> None:
        """记录不经过主循环的动作（回费速度）：不改变当前动作显示。"""
        with self._lock:
            order = None
            if getattr(action, "enabled", True):
                self._order += 1
                order = self._order
            self._results[int(index)] = {
                "ok": True,
                "note": note,
                "order": order,
                "action": action,
            }
            self._result_dirty = True
            self._dirty = True

    def finish(self, ok: bool = True) -> None:
        with self._lock:
            self._finished = True
            self._ok = bool(ok)
            if not ok:
                self._status = "失败"
            elif self._failures:
                self._status = "完成（有失败动作）"
            else:
                self._status = "完成"
            self._current_text = ""
            self._current_index = None
            self._dirty = True
            self._map_dirty = True

    def stop(self) -> None:
        if not self.enabled:
            self._write_json()
            return
        self._stop = True
        if self._thread is not None:
            self._thread.join(timeout=8.0)
        # 不管 Excel 面板是否打开，都写一份最新 JSON，避免 run_status.json
        # 一直停在上一局（用户排查时会看它）。
        self._write_json()

    # ---------- 后台线程 ----------
    def _run(self) -> None:
        # 延迟导入：绝不能在主线程 import，否则 WinRT 抓屏会失败。
        try:
            import pythoncom
            # 先初始化 STA，再导入 win32com。某些 pywin32 版本在导入
            # 时就会触碰 COM，顺序反了会把线程留在错误的 apartment。
            try:
                pythoncom.CoInitializeEx(
                    pythoncom.COINIT_APARTMENTTHREADED)
            except Exception:
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    pass
            import win32com.client as win32_client
        except Exception as e:
            self.start_error = f"pywin32/COM 不可用：{e!r}"
            print(f"[excel] {self.start_error}，面板关闭")
            self._excel_ok = False
            self._ready.set()
            self._write_json()
            return
        self._win32_client = win32_client
        try:
            self._open()
            self._excel_ok = True
            self._ready.set()
            while not self._stop:
                self._process_sync()
                self._poll_runtime()
                with self._lock:
                    dirty = self._dirty
                    self._dirty = False
                    self._map_dirty = False
                wrote = False
                if dirty:
                    try:
                        wrote |= self._flush()
                    except Exception as e:
                        if (self._is_wrong_thread_error(e)
                                or self._is_excel_gone_error(e)):
                            # 公寓/权限不一致，或 Excel 已被关闭：重试永远
                            # 不会成功，直接退出面板线程，后续只写 JSON。
                            raise
                        # 用户在 Excel 里编辑单元格时，COM 会短暂拒绝调用
                        # （0x80010001）。不能因此让面板线程退出；保留
                        # dirty，下个循环继续写。同一错误最多 5s 打一行。
                        self._warn_throttled(
                            "[excel] 写状态被 Excel 拒绝，稍后重试", e)
                        with self._lock:
                            self._dirty = True
                # 每轮都检查地图签名。状态没变时不碰 Excel（不会闪）；
                # 变了立即画，不依赖 action_start/action_done 有没有置脏。
                try:
                    wrote |= self._flush_map()
                except Exception as e:
                    if (self._is_wrong_thread_error(e)
                            or self._is_excel_gone_error(e)):
                        raise
                    self._warn_throttled(
                        "[excel] 地图刷新失败（稍后重试）", e)
                    with self._lock:
                        self._map_dirty = True
                # 只有结果变化才需要保存；状态/地图刷新不落盘。
                with self._lock:
                    result_dirty = self._result_dirty
                if result_dirty and time.time() - self._result_saved_at >= 1.0:
                    if self._save_now():
                        with self._lock:
                            self._result_dirty = False
                            self._result_saved_at = time.time()
                if self._save_request.is_set():
                    self._save_request.clear()
                    self._save_ok = False
                    for attempt in range(3):
                        try:
                            if self._wb is not None:
                                self._wb.Save()
                                self._save_ok = True
                                print("[excel] 已保存工作簿（重载操作前）")
                            break
                        except Exception as e:
                            print(f"[excel] 保存工作簿失败"
                                  f"（{attempt + 1}/3）：{e!r}")
                            time.sleep(0.2)
                    self._save_done.set()
                if self._front_request.is_set():
                    self._front_request.clear()
                    ok = False
                    try:
                        if self._app is not None and self._wb is not None:
                            self._app.Visible = True
                            self._wb.Activate()
                            import win32gui
                            hwnd = int(self._app.Hwnd)
                            win32gui.ShowWindow(hwnd, 9)   # SW_RESTORE
                            # 直接抢可能被 Windows 前台锁挡住；附到当前前台线程
                            # 再抢一次，绕过前台锁（对齐 axis_console._focus_window）。
                            try:
                                win32gui.SetForegroundWindow(hwnd)
                            except Exception:
                                pass
                            if win32gui.GetForegroundWindow() != hwnd:
                                import ctypes
                                u32 = ctypes.windll.user32
                                k32 = ctypes.windll.kernel32
                                fg = int(u32.GetForegroundWindow() or 0)
                                tid_fg = ctypes.c_ulong()
                                u32.GetWindowThreadProcessId(
                                    fg, ctypes.byref(tid_fg))
                                tid_me = k32.GetCurrentThreadId()
                                u32.AttachThreadInput(tid_me, tid_fg.value, 1)
                                try:
                                    u32.SetForegroundWindow(hwnd)
                                finally:
                                    u32.AttachThreadInput(
                                        tid_me, tid_fg.value, 0)
                            ok = (win32gui.GetForegroundWindow() == hwnd)
                    except Exception as e:
                        print(f"[excel] Excel 前置失败：{e!r}")
                    self._front_ok = ok
                    if not ok:
                        print("[excel] Excel 前置失败：窗口没能切到前台")
                    self._front_done.set()
                # 同步请求会 set _wake，立即醒来；平时 50ms 刷新状态。
                self._wake.wait(0.05)
                self._wake.clear()
            self._poll_runtime()
            with self._lock:
                self._dirty = False
                self._map_dirty = False
            self._flush()
            # 最后再强制画一次；COM 忙时重试，避免收尾时地图少一个标记。
            for attempt in range(3):
                try:
                    self._flush_map(force=True)
                    break
                except Exception as e:
                    print(f"[excel] 收尾地图刷新失败"
                          f"（{attempt + 1}/3）：{e!r}")
                    time.sleep(0.15)
            self._save_now()
            print("[excel] 已保存运行面板（Excel 窗口保持打开）")
        except Exception as e:
            self._excel_ok = False
            self.start_error = f"{type(e).__name__}: {e}"
            if self._is_excel_gone_error(e):
                print("[excel] Excel 已关闭，面板停止写入；"
                      "控制台仍可用，run_status.json 保留最后状态")
            else:
                print(f"[excel] 面板异常，后续只写 JSON：{e!r}")
                traceback.print_exc()
            # 打开到一半失败时，把自己开的只读/半开窗口关掉，免得挡住
            # 用户。连接的是用户已经打开的 Excel 时绝对不能关。
            if self._attached:
                print("[excel] 保留你打开的 Excel，不关闭")
            else:
                for _ in range(3):
                    try:
                        if self._wb is not None:
                            self._wb.Close(SaveChanges=False)
                        if self._app is not None:
                            self._app.Quit()
                        break
                    except Exception:
                        time.sleep(0.4)
            self._wb = None
            self._app = None
            self._write_json()
        finally:
            self._ready.set()
            if self.quit_on_stop and not self._attached:
                try:
                    if self._wb is not None:
                        self._wb.Close(SaveChanges=True)
                    if self._app is not None:
                        self._app.Quit()
                except Exception:
                    pass
            self._wb = None
            self._app = None
            self._attached = False
            self._ws_rec = None
            self._ws_mid = None
            self._ws_map = None
            try:
                if pythoncom is not None:
                    pythoncom.CoUninitialize()
            except Exception:
                pass

    def _save_now(self) -> bool:
        """立即保存工作簿；Excel 忙时短暂重试。"""
        if self._wb is None:
            return False
        for attempt in range(3):
            try:
                self._wb.Save()
                return True
            except Exception as e:
                print(f"[excel] 立即保存失败（{attempt + 1}/3）：{e!r}")
                time.sleep(0.05)
        return False

    def _lock_path(self) -> str:
        return os.path.join(
            os.path.dirname(self.report_path),
            "~$" + os.path.basename(self.report_path))

    @staticmethod
    def _norm_path(path) -> str:
        return os.path.normcase(os.path.abspath(str(path)))

    def _find_open_workbook(self, app, target: str):
        try:
            count = int(app.Workbooks.Count)
        except Exception:
            return None
        for i in range(1, count + 1):
            try:
                wb = app.Workbooks(i)
                if self._norm_path(wb.FullName) == target:
                    return wb
            except Exception:
                continue
        return None

    def _coerce_workbook(self, disp, target: str):
        """ROT 里拿到的可能是 Workbook、Application 或别的对象。"""
        try:
            if self._norm_path(disp.FullName) == target:
                return disp
        except Exception:
            pass
        try:
            app = disp.Application
            wb = self._find_open_workbook(app, target)
            if wb is not None:
                return wb
        except Exception:
            pass
        try:
            wb = self._find_open_workbook(disp, target)
            if wb is not None:
                return wb
        except Exception:
            pass
        return None

    @staticmethod
    def _is_wrong_thread_error(exc: Exception) -> bool:
        """判断是否是 Excel 的 RPC_E_WRONG_THREAD (0x8001010E)。"""
        code = getattr(exc, "hresult", None)
        if code is None:
            for arg in getattr(exc, "args", ()):
                if isinstance(arg, int):
                    code = arg
                    break
        try:
            return (int(code) & 0xFFFFFFFF) == 0x8001010E
        except (TypeError, ValueError):
            return False

    @staticmethod
    def _is_excel_gone_error(exc: Exception) -> bool:
        """Excel 已被关闭 / COM 对象失效：重试不会成功，必须退出面板线程。

        0x800A01A8 OLE Object required（关掉 Excel 后 _wb/_ws_* 变成悬空
        代理，就是控制台一直刷的那条）、0x800706BA RPC 服务器不可用、
        0x80010108 对象已断开连接。
        """
        code = getattr(exc, "hresult", None)
        if code is None:
            for arg in getattr(exc, "args", ()):
                if isinstance(arg, int):
                    code = arg
                    break
        try:
            return (int(code) & 0xFFFFFFFF) in (
                0x800A01A8, 0x800706BA, 0x80010108)
        except (TypeError, ValueError):
            return False

    def _warn_throttled(self, prefix: str, exc: Exception) -> None:
        """可恢复的 COM 错误最多每 5 秒打印一行，避免把控制台刷没。"""
        now = time.time()
        if now - self._last_com_warn_at < 5.0:
            return
        self._last_com_warn_at = now
        print(f"{prefix}：{exc!r}")

    def _use_open_workbook(self, app, wb) -> None:
        if bool(wb.ReadOnly):
            raise RuntimeError(
                "轴文件在 Excel 里是只读打开；请关闭后重新打开并保存")
        try:
            saved = bool(wb.Saved)
        except Exception:
            saved = True
        if not saved:
            raise RuntimeError(
                "轴文件在 Excel 里有未保存修改；请先在 Excel 里保存，再重新运行")
        # 能拿到 wb 不代表这个 COM 对象在本线程可用。以前正是在这里
        # “连接成功”后，所有写入才报 RPC_E_WRONG_THREAD。先实际读一下
        # 工作表并调用 Calculate，确认能控制它。
        last_exc = None
        for _ in range(3):
            try:
                _ = int(wb.Worksheets.Count)
                _ = wb.Worksheets(1).Range("A1").Value
                _ = bool(app.Visible)
                app.Calculate()
                last_exc = None
                break
            except Exception as exc:
                last_exc = exc
                time.sleep(0.2)
        if last_exc is not None:
            raise RuntimeError(
                "检测到轴文件已在 Excel 中打开，但本程序的 COM 线程"
                "无法控制这个 Excel。请先在 Excel 里保存并完全关闭文件"
                "（包括任务管理器里残留的 EXCEL.EXE），再重新 load。"
                f"原始错误：{last_exc!r}") from last_exc
        self._app = app
        self._wb = wb
        self._attached = True

    def _attach_open_workbook(self) -> None:
        """连接用户已经打开的 Excel，而不是另开一个实例。

        先试 GetActiveObject，再扫 ROT（工作表窗口标题/全路径）。
        不走 GetObject(路径)：文件已被占用时它会尝试再开一个实例并
        弹只读对话框，可能卡死。
        """
        target = self._norm_path(self.report_path)
        target_base = os.path.basename(target)

        try:
            app = self._win32_client.GetActiveObject("Excel.Application")
            wb = self._find_open_workbook(app, target)
            if wb is not None:
                self._use_open_workbook(app, wb)
                return
        except Exception as e:
            print(f"[excel] GetActiveObject 连接失败：{e!r}")

        try:
            import pythoncom
            ctx = pythoncom.CreateBindCtx(0)
            rot = pythoncom.GetRunningObjectTable()
            for moniker in rot.EnumRunning().Next(100):
                try:
                    name = moniker.GetDisplayName(ctx, None) or ""
                except Exception:
                    continue
                low = os.path.normcase(name)
                if target not in low and not (
                        target_base and low.endswith(target_base)):
                    continue
                try:
                    obj = rot.GetObject(moniker)
                    # ROT 给的是 IUnknown；先 QI 到 IDispatch，win32com
                    # 的 Dispatch 才能直接把它当 Workbook 用。
                    try:
                        disp = self._win32_client.Dispatch(
                            obj.QueryInterface(pythoncom.IID_IDispatch))
                    except Exception:
                        disp = self._win32_client.Dispatch(obj)
                except Exception as e:
                    print(f"[excel] ROT 绑定失败：{e!r}")
                    continue
                wb = self._coerce_workbook(disp, target)
                if wb is not None:
                    self._use_open_workbook(wb.Application, wb)
                    return
        except Exception as e:
            print(f"[excel] ROT 连接失败：{e!r}")

        raise RuntimeError(
            "检测到轴文件已在 Excel 中打开，但本程序连不上这个 Excel。"
            "请先在 Excel 里保存并完全关闭文件（包括任务管理器里残留的"
            " EXCEL.EXE），再重新 load；如果仍失败，请确认 Excel 和"
            "本工具权限相同。")

    def _open(self) -> None:
        if self.in_place and os.path.exists(self._lock_path()):
            print("[excel] 检测到轴文件已在 Excel 中打开，尝试直接连接…")
            self._attach_open_workbook()
            print("[excel] 已连接现有 Excel，不再重复打开")
        else:
            self._app = self._win32_client.DispatchEx("Excel.Application")
            self._app.Visible = True
            self._app.DisplayAlerts = False
            # 第二个位置参数是 UpdateLinks=0，避免打开时弹外部链接提示。
            self._wb = self._app.Workbooks.Open(
                os.path.abspath(self.report_path), 0)
        if bool(self._wb.ReadOnly):
            raise RuntimeError(
                "Excel 以只读方式打开了文件（可能正被另一个 Excel 占用）；"
                "请先保存并关闭它，再重新运行")
        # Excel 刚打开时偶尔会拒绝一次 COM 调用（0x800AC472）。
        time.sleep(0.35)
        if self.in_place:
            # 原地模式沿用用户文件的列：不强行改成 sample 的 B~H / I。
            self._cols = {
                "当前执行": (int(self.axis.columns["当前执行"]) + 1
                             if "当前执行" in self.axis.columns else None),
                "运行结果": (int(self.axis.columns["运行结果"]) + 1
                             if "运行结果" in self.axis.columns else None),
            }
        else:
            self._cols = {"当前执行": 1, "是否运行": 10, "运行结果": 11}

        # 第 1 页：作战记录（保持 sample 模板的列和设置区）
        self._ws_rec = self._find_sheet("作战记录") or self._wb.Worksheets(1)
        if str(self._ws_rec.Name) != "作战记录":
            self._ws_rec.Name = "作战记录"
        self._prepare_rec()

        # 第 2 页：运行信息。
        # 兼容旧模板：旧文件里可能叫“地图”“地图与中转站”或“中转站”。
        mid = (self._find_sheet("运行信息")
               or self._find_sheet("地图")
               or self._find_sheet("地图与中转站")
               or self._find_sheet("中转站"))
        if mid is None:
            mid = self._wb.Worksheets.Add(
                After=self._wb.Worksheets(self._wb.Worksheets.Count))
        mid.Name = "运行信息"
        # 不清空整页：只重写状态区。
        self._setup_status(mid)
        self._ws_mid = mid

        # 第 3 页：地图示意（独立一页）。
        map_ws = self._find_sheet("地图示意")
        if map_ws is None:
            legacy = self._find_sheet("地图")
            if legacy is not None and str(legacy.Name) != "运行信息":
                map_ws = legacy
                map_ws.Name = "地图示意"
            else:
                map_ws = self._wb.Worksheets.Add(
                    None, self._wb.Worksheets(self._wb.Worksheets.Count))
                map_ws.Name = "地图示意"
        try:
            # 固定第三页：作战记录 / 运行信息 / 地图示意。
            map_ws.Move(None, mid)
        except Exception as exc:
            print(f"[excel] 地图页排序失败（忽略）：{exc!r}")
        self._ws_map = map_ws

        # 运行汇总页已取消；旧文件里如果还留着，删掉，避免一页过期数据。
        old_sum = self._find_sheet("运行汇总")
        if old_sum is not None:
            try:
                old_sum.Delete()
                print("[excel] 已移除旧的“运行汇总”页")
            except Exception as e:
                print(f"[excel] 移除旧汇总页失败（忽略）：{e!r}")

        try:
            self._draw_map_static()
        except Exception as e:
            print(f"[excel] 地图初始化失败（面板其余部分继续）：{e!r}")
        self._flush(force=True)
        try:
            self._flush_map(force=True)
        except Exception:
            pass
        self._save_now()

    def _find_sheet(self, name: str):
        for i in range(1, self._wb.Worksheets.Count + 1):
            ws = self._wb.Worksheets(i)
            if str(ws.Name) == name:
                return ws
        return None

    def _ensure_axis_mark_setting(self, ws) -> None:
        """旧轴设置区没有“打轴时暂停”时补一行，默认 TRUE。"""
        settings_col = None
        cols = max(int(ws.UsedRange.Columns.Count), 14)
        for c in range(1, cols + 1):
            if str(ws.Cells(1, c).Value or "").strip() == "设置":
                settings_col = c
                break
        if settings_col is None:
            return
        last = 1
        rows = max(int(ws.UsedRange.Rows.Count), 2)
        for r in range(2, rows + 1):
            key = str(ws.Cells(r, settings_col).Value or "").strip()
            if key == "打轴时暂停":
                return
            if key:
                last = r
        ws.Cells(last + 1, settings_col).Value = "打轴时暂停"
        ws.Cells(last + 1, settings_col + 1).Value = True

    def _prepare_rec(self) -> None:
        """第 1 页。

        普通模式：按 sample 版式清空 A-I / L-N，重写动作和设置。
        原地模式：只补空表头和缺失的“运行结果”列，绝不清空用户填的操作行。
        """
        ws = self._ws_rec
        headers = ["当前执行", "帧号1", "帧号2", "操作", "单位", "坐标",
                   "朝向", "简称", "备注", "是否运行", "运行结果"]

        if self.in_place:
            # 只补输入列的空表头；不动“运行结果”，它由下面单独定位。
            for c, text in enumerate(headers[:-1], start=1):
                if not str(ws.Cells(1, c).Value or "").strip():
                    ws.Cells(1, c).Value = text
            # 旧模板的“干员”表头统一改成“单位”（F5 打轴和回费速度操作都用这列）。
            for c in range(1, int(ws.UsedRange.Columns.Count) + 1):
                if str(ws.Cells(1, c).Value or "").strip() == "干员":
                    ws.Cells(1, c).Value = "单位"
            # 表头字体对齐 A1（用户看到“运行结果”字体不一致时自动修）。
            try:
                base_font = ws.Cells(1, 1).Font
                for c in range(1, len(headers) + 1):
                    f = ws.Cells(1, c).Font
                    f.Name = base_font.Name
                    f.Size = base_font.Size
                    f.Bold = base_font.Bold
            except Exception:
                pass
            # 设置区：已有“设置”就沿用；没有就在 L1 补一个表头。
            has_settings = False
            used_cols = int(ws.UsedRange.Columns.Count)
            for c in range(1, max(used_cols, 14) + 1):
                if str(ws.Cells(1, c).Value or "").strip() == "设置":
                    has_settings = True
                    break
            if not has_settings:
                self._cell(ws, 1, 12, "设置", bold=True, border=True)
            self._ensure_axis_mark_setting(ws)
            # 没有“运行结果”列就补在最后输入列后面。
            if self._cols.get("运行结果") is None:
                if self.axis.columns:
                    col = max(self.axis.columns.values()) + 2
                else:
                    col = 9
                self._cell(ws, 1, col, "运行结果", bold=True, border=True)
                self._cols["运行结果"] = col

            # 新一局开始：清掉上一次留下的“运行结果”和箭头，但保留
            # 用户填的 B~H。否则面板在 F6 才打开时，旧结果会一直挂着。
            result_col = self._cols.get("运行结果")
            if result_col:
                last = max(int(ws.UsedRange.Rows.Count), 2)
                ws.Range(ws.Cells(2, result_col),
                         ws.Cells(last, result_col)).ClearContents()
            arrow_col = self._cols.get("当前执行")
            if arrow_col:
                for a in self.axis.actions:
                    if a.excel_row:
                        ws.Cells(int(a.excel_row), arrow_col).Value = ""
            return

        for c, text in enumerate(headers, start=1):
            ws.Cells(1, c).Value = text
        ws.Cells(1, 12).Value = "设置"
        last = max(int(ws.UsedRange.Rows.Count), 12)
        # 只清内容，保留 sample 的行高/边框/字体
        ws.Range(ws.Cells(2, 1), ws.Cells(last, 11)).ClearContents()
        ws.Range(ws.Cells(2, 12), ws.Cells(last, 14)).ClearContents()

        for idx, a in enumerate(self.axis.actions):
            r = 2 + idx
            ws.Cells(r, 2).Value = a.frame
            ws.Cells(r, 3).Value = a.frame2 if a.frame2 is not None else ""
            ws.Cells(r, 4).Value = a.kind
            ws.Cells(r, 5).Value = a.oper or ""
            ws.Cells(r, 6).Value = a.pos or ""
            ws.Cells(r, 7).Value = a.direction or ""
            ws.Cells(r, 8).Value = a.alias or ""
            ws.Cells(r, 9).Value = a.note or ""
            ws.Cells(r, 10).Value = bool(a.enabled)

        ws.Cells(2, 12).Value = "关卡代号"
        # “1-7”直接写会被 Excel 当成日期（46029），强制文本
        ws.Cells(2, 13).NumberFormat = "@"
        ws.Cells(2, 13).Value = self.axis.map_code or ""
        ws.Cells(3, 12).Value = "回费速度"
        raw = self.axis.settings.get("回费速度")
        if raw is None:
            raw = self.axis.settings.get("可出现回费速度")
        if isinstance(raw, (list, tuple)):
            values = list(raw)
        elif raw is None:
            values = []
        else:
            values = [raw]
        if not values:
            declared = self.rate_info.get("declared") or []
            base = float(self.rate_info.get("base", 100.0))
            values = ([x / 100.0 for x in declared] if declared
                      else [base / 100.0])
        for k, v in enumerate(values):
            ws.Cells(3, 13 + k).Value = v
        # 新设置区同步到副本报表。
        for r, key, default in (
                (4, "默认2倍速", False),
                (5, "开始前暂停", False),
                (6, "结束后暂停", False),
                (7, "1x操作提前量", 4),
                (8, "2x操作提前量", 5),
                (9, "打轴时暂停", True),
                (10, "低回费盲步进", True)):
            ws.Cells(r, 12).Value = key
            value = self.axis.settings.get(key)
            if isinstance(default, bool):
                ws.Cells(r, 13).Value = _to_bool(value, default)
            else:
                ws.Cells(r, 13).Value = default if value is None else value

    def _setup_status(self, ws) -> None:
        # 行 1~3 保持 A 标签 / B 值。
        for r, (label, value) in enumerate(STATUS_ROWS, start=1):
            self._cell(ws, r, 1, label, bold=True)
            self._cell(ws, r, 2, value or "")
        # 行 5 起改成 A 标签 / B 值；先清掉旧版残留。
        for r in range(4, 18):
            for c in range(1, 3):
                ws.Cells(r, c).Value = ""
        for r, label in STATUS_FIELDS:
            self._cell(ws, r, 1, label, bold=True)
        # 文本字段：帧号/时间/费用帧/回费/倍速/选中。
        for r in (5, 6, 7, 8, 9, 13):
            ws.Cells(r, 2).NumberFormat = "@"

        runtime = self._counter
        cur = None
        if self._executor is not None and hasattr(self._executor,
                                                     "current_frame"):
            try:
                cur = int(self._executor.current_frame())
            except Exception:
                cur = None
        if cur is None and runtime is not None:
            try:
                cur = int(runtime.frame)
            except Exception:
                cur = None
        ws.Cells(5, 2).Value = "" if cur is None else str(cur)
        ws.Cells(6, 2).Value = "" if cur is None else _frame_time_text(cur)
        last = getattr(runtime, "last", None) if runtime is not None else None
        ws.Cells(7, 2).Value = _cost_frame_text(last)
        ws.Cells(8, 2).Value = self._realtime_rate_text(False)
        if last is not None:
            bullet = bool(getattr(last, "bullet", False))
            m = float(getattr(last, "m", 1.0) or 1.0)
            ws.Cells(9, 2).Value = ("0.2X" if bullet
                                    else ("2X" if m >= 1.5 else "1X"))

    # ---------- 状态读取 ----------
    def _compute_runtime_status(self) -> dict:
        """计算运行信息页的实时字段；不碰 Excel，控制台也能调用。"""
        with self._compute_lock:
            runtime = self._counter
            ex = self._executor
            frame_text = ""
            time_text = ""
            cost_frame_text = ""
            rate_text = self._realtime_rate_text(False)
            status = None
            error = ""
            speed_text = ""
            pause_text = ""
            selected_text = ""

            if runtime is not None:
                try:
                    # 当前游戏帧 cur（含 offset/盲步进预测）。
                    cur = (int(ex.current_frame())
                           if ex is not None and hasattr(ex, "current_frame")
                           else int(runtime.frame))
                    frame_text = str(cur)
                    time_text = _frame_time_text(cur)
                    cost_frame_text = _cost_frame_text(runtime.last)
                    # 以费用条状态机当前判定为准；还没读到画面时退回设置值。
                    neg = bool(getattr(runtime.counter, "neg",
                                       getattr(runtime.last, "neg", False)))
                    rate_text = self._realtime_rate_text(neg)
                    error = getattr(runtime.counter, "spec_error", "") or ""
                except Exception:
                    pass

            m = 1.0
            bullet = False
            if runtime is not None and runtime.last is not None:
                try:
                    m = float(getattr(runtime.last, "m", 1.0) or 1.0)
                except (TypeError, ValueError):
                    m = 1.0
                bullet = bool(getattr(runtime.last, "bullet", False))
            hud = {}
            in_battle = bool(getattr(runtime, "in_battle", False))
            # 只在关卡内做 HUD OCR；主菜单/加载/结算不检测。
            if (in_battle and ex is not None
                    and getattr(ex, "stream", None) is not None):
                now = time.time()
                if now - self._hud_read_at >= 0.3:
                    self._hud_read_at = now
                    try:
                        img = ex.stream.latest()
                        if img is not None:
                            self._ocr.submit(img)
                    except Exception:
                        pass
                hud = self._ocr.values()

            cost_text = "" if hud.get("cost") is None else str(hud["cost"])
            deploy_text = ("" if hud.get("deploy_left") is None
                           else str(hud["deploy_left"]))
            stat = {k: ("" if hud.get(k) is None else str(hud[k]))
                    for k in ("attack", "defense", "resist", "block")}

            with self._lock:
                finished = self._finished
                ok = self._ok
                failures = self._failures

            if runtime is None:
                status = "准备中"
            else:
                # 子弹时间固定 0.2×，否则按识别出的 m 显示 1X/2X。
                speed_text = ("0.2X" if bullet
                              else ("2X" if m >= 1.5 else "1X"))
                if bullet:
                    name = hud.get("name") or ""
                    if not name and ex is not None:
                        name = getattr(ex, "selected_oper", None)
                    if not name and ex is not None:
                        name = getattr(ex, "selected_name", None)
                    if ex is not None:
                        ex.selected_name = name or None
                    selected_text = name or "?"
                else:
                    if ex is not None:
                        selected_text = ""
                        ex.selected_name = None
                    stat = {k: "" for k in stat}
                status = getattr(ex, "phase", None) or "运行中"
                if finished:
                    if not ok:
                        status = "轴完成（失败）"
                    elif failures:
                        status = "轴完成（有失败动作）"
                    else:
                        status = "轴完成"

            return {
                "status": status or "准备中",
                "frame": frame_text,
                "time": time_text,
                "cost_frame": cost_frame_text,
                "rate": rate_text,
                "speed": speed_text,
                "pause": pause_text,
                "selected": selected_text,
                "cost": cost_text,
                "deploy_left": deploy_text,
                "attack": stat["attack"],
                "defense": stat["defense"],
                "resist": stat["resist"],
                "block": stat["block"],
                "error": error,
            }

    def status_snapshot(self) -> dict:
        """控制台 status 用：实时字段 + 当前操作/已有报错。"""
        data = self._compute_runtime_status()
        with self._lock:
            data["current"] = self._current_text
            if self._error:
                data["error"] = self._error
        return data

    def _poll_runtime(self) -> None:
        data = self._compute_runtime_status()
        with self._lock:
            self._status = data["status"]
            self._frame_text = data["frame"]
            self._time_text = data["time"]
            self._cost_frame_text = data["cost_frame"]
            self._rate_text = data["rate"]
            self._speed_text = data["speed"]
            self._pause_text = data["pause"]
            self._cost_text = data["cost"]
            self._deploy_left_text = data["deploy_left"]
            self._attack_text = data["attack"]
            self._defense_text = data["defense"]
            self._resist_text = data["resist"]
            self._block_text = data["block"]
            self._selected_text = data["selected"]
            if data["error"] and not self._error:
                self._error = data["error"]
            self._dirty = True
        self._refresh_enabled()

    def _process_sync(self) -> None:
        """处理主线程的同步屏障请求（在 Excel 线程里执行）。"""
        with self._sync_lock:
            kind = self._sync_kind
            index = self._sync_index
            ok = self._sync_ok
            payload = self._sync_payload
        if kind is None or index is None:
            return
        result = None
        try:
            if kind == "add_row":
                ws = self._ws_rec
                used = int(ws.UsedRange.Rows.Count) + 1
                last = 1
                for r in range(2, used + 1):
                    b = ws.Cells(r, 2).Value
                    d = ws.Cells(r, 4).Value
                    if (b not in (None, "")) or (d not in (None, "")):
                        last = r
                new_row = last + 1
                # 新行默认套的是工作表默认样式（微软雅黑 10 + 边框），和模板
                # 数据行（宋体 11、无边框）不一样，一眼就能看出“这行是程序加的”。
                # 只把**格式**从数据行复制下来，不复制内容。
                # 样式源固定取第 2 行（第一条记录），不要取 last：last 可能
                # 本身就是之前 F5 加的行，风格会一路传下去。
                style_row = 2 if ws.Cells(2, 2).Value not in (None, "") else last
                if style_row >= 2:
                    try:
                        src = ws.Range(ws.Cells(style_row, 1),
                                       ws.Cells(style_row, 11))
                        dst = ws.Range(ws.Cells(new_row, 1),
                                       ws.Cells(new_row, 11))
                        src.Copy()
                        dst.PasteSpecial(-4122)          # xlPasteFormats
                    except Exception as exc:
                        print(f"[excel] 复制行格式失败（忽略）：{exc!r}")
                    finally:
                        try:
                            if self._app is not None:
                                self._app.CutCopyMode = False
                        except Exception:
                            pass
                ws.Cells(new_row, 2).Value = int(payload)
                # 不在这里立即保存：F5 要马上暂停并前置 Excel，保存会拖慢。
                # 之后的结果定时保存 / F6 保存会一并落盘。
                result = new_row
            elif kind == "sort":
                order_rows, last_row = payload
                ws = self._ws_rec
                ws.Columns(12).Insert()
                try:
                    rank = {row: i + 1 for i, row in enumerate(order_rows)}
                    for r in range(2, int(last_row) + 1):
                        ws.Cells(r, 12).Value = (
                            rank.get(r, len(order_rows) + r))
                    ws.Range(ws.Cells(2, 1),
                             ws.Cells(int(last_row), 12)).Sort(
                        Key1=ws.Cells(2, 12), Order1=1, Header=2)
                finally:
                    ws.Columns(12).Delete()
                result = True
            elif kind == "formulas":
                for item in payload:
                    ws = self._wb.Worksheets(item["sheet"])
                    ws.Cells(int(item["row"]), int(item["col"])).Formula = \
                        item["formula"]
                self._app.Calculate()
                result = True
            else:
                if not (0 <= index < len(self.axis.actions)):
                    raise IndexError(f"index={index}")
                action = self.axis.actions[index]
                row = (int(action.excel_row)
                       if self.in_place and action.excel_row else 2 + index)
                if kind == "read":
                    self._app.Calculate()
                    value = self._ws_rec.Cells(row, 10).Value
                    result = _to_bool(value, True)
                elif kind == "commit":
                    if ok is None:
                        self._ws_rec.Cells(row, 11).Value = None
                    else:
                        self._ws_rec.Cells(row, 11).Value = bool(ok)
                    self._app.Calculate()
                    result = True
        except Exception as exc:
            print(f"[excel] 同步屏障失败（{kind} 行 {index}）：{exc!r}")
            result = None
        finally:
            with self._sync_lock:
                self._sync_kind = None
                self._sync_index = None
                self._sync_result = result
            self._sync_done.set()

    def _refresh_enabled(self) -> None:
        """定期读 J 列“是否运行”，让 Excel 公式/手工修改在运行中生效。"""
        if self._ws_rec is None or not self.axis.actions:
            return
        now = time.time()
        if now - self._enabled_read_at < 0.3:
            return
        self._enabled_read_at = now
        try:
            rows = []
            for idx, a in enumerate(self.axis.actions):
                row = (int(a.excel_row) if self.in_place and a.excel_row
                       else 2 + idx)
                rows.append((idx, a, row))
            last = max(r for _, _, r in rows)
            data = self._ws_rec.Range(
                self._ws_rec.Cells(2, 10),
                self._ws_rec.Cells(last, 10)).Value
            if not isinstance(data, tuple):
                data = ((data,),)
            for idx, a, row in rows:
                pos = row - 2
                if 0 <= pos < len(data):
                    cell = data[pos]
                    value = cell[0] if isinstance(cell, tuple) else cell
                    a.enabled = _to_bool(value, True)
        except Exception:
            pass

    # ---------- 写 Excel ----------
    def _cell(self, ws, row: int, col: int, value, *,
              fill: str | None = None, bold: bool = False,
              size: int | None = None, color: str | None = None,
              border: bool = False, border_color: str | None = None):
        cell = ws.Cells(row, col)
        cell.Value = value
        if fill:
            cell.Interior.Color = _excel_color(fill)
        if bold or size or color:
            cell.Font.Bold = bool(bold)
            if size:
                cell.Font.Size = size
            if color:
                cell.Font.Color = _excel_color(color)
        if border:
            # Excel COM 偶尔会在设置 Borders.Color 时抛 DISP_E_EXCEPTION；
            # 边框只是装饰，失败不能让整个面板挂掉。
            try:
                cell.Borders.LineStyle = 1
                cell.Borders.Weight = 2
                if border_color:
                    cell.Borders.Color = _excel_color(border_color)
            except Exception:
                pass
        return cell

    def _flush(self, force: bool = False) -> bool:
        if self._ws_mid is None or self._ws_rec is None:
            return False
        with self._lock:
            status = self._status
            current = self._current_text
            error = self._error
            frame = self._frame_text
            time_text = self._time_text
            cost_frame_text = self._cost_frame_text
            rate = self._rate_text
            speed = self._speed_text
            pause = self._pause_text
            selected = self._selected_text
            cost = self._cost_text
            deploy_left = self._deploy_left_text
            attack = self._attack_text
            defense = self._defense_text
            resist = self._resist_text
            block = self._block_text
            index = self._current_index
            results = dict(self._results)
            finished = self._finished

        def _num(value):
            s = str(value or "").strip()
            if not s:
                return ""
            try:
                return int(s)
            except ValueError:
                return s

        wrote = False
        if force or self._written.get("B1") != status:
            self._cell(self._ws_mid, 1, 2, status)
            self._written["B1"] = status
            wrote = True
        if force or self._written.get("B2") != current:
            self._cell(self._ws_mid, 2, 2, current)
            self._written["B2"] = current
            wrote = True
        if force or self._written.get("B3") != error:
            self._cell(self._ws_mid, 3, 2, error)
            self._written["B3"] = error
            wrote = True
        # 行 5 起：A 列标签，B 列值。
        for row, label in STATUS_FIELDS:
            key = f"L{row}"
            if force or self._written.get(key) != label:
                self._cell(self._ws_mid, row, 1, label, bold=True)
                self._written[key] = label
                wrote = True
        for key, row, value in (
                ("B5", 5, frame), ("B6", 6, time_text),
                ("B7", 7, cost_frame_text), ("B8", 8, rate),
                ("B9", 9, speed), ("B10", 10, _num(cost)),
                ("B11", 11, _num(deploy_left)), ("B13", 13, selected),
                ("B14", 14, _num(attack)), ("B15", 15, _num(defense)),
                ("B16", 16, _num(resist)), ("B17", 17, _num(block))):
            if force or self._written.get(key) != value:
                self._cell(self._ws_mid, row, 2, value)
                self._written[key] = value
                wrote = True

        # 当前执行 = 运行顺序（1 基）；运行结果 = 布尔值。
        order_col = self._cols.get("当前执行")
        result_col = self._cols.get("运行结果")

        def rec_row(i: int) -> int:
            if 0 <= i < len(self.axis.actions):
                a = self.axis.actions[i]
                if self.in_place and a.excel_row:
                    return int(a.excel_row)
            return 2 + i

        if result_col:
            for idx, r in results.items():
                if not (0 <= idx < len(self.axis.actions)):
                    continue
                row = rec_row(idx)
                order = r.get("order")
                if order_col:
                    order_value = "" if order is None else int(order)
                    key = f"{order_col}:{row}"
                    if force or self._written.get(key) != repr(order_value):
                        self._ws_rec.Cells(row, order_col).Value = order_value
                        self._written[key] = repr(order_value)
                        wrote = True
                if r.get("note") == "无操作":
                    value = "无操作"
                elif r.get("ok") is None:
                    value = ""
                else:
                    value = bool(r["ok"])
                key = f"{result_col}:{row}"
                if force or self._written.get(key) != repr(value):
                    self._ws_rec.Cells(row, result_col).Value = value
                    self._written[key] = repr(value)
                    wrote = True
        return wrote

    def _ensure_label(self, row: int, text: str) -> None:
        cell = self._ws_mid.Cells(row, 1)
        if not cell.Value:
            cell.Value = text

    # ---------- 地图 ----------
    def _draw_map_static(self) -> None:
        ws = self._ws_map
        if ws is None:
            return
        self._terrain.clear()
        self._marked.clear()
        self._map_signature = None
        data = self.axis.map_data
        # 兼容旧模板：地图标题可能在第 9 或第 10 行。
        title_row = MAP_TITLE_ROW
        try:
            last = min(int(ws.UsedRange.Rows.Count), 20)
            for r in range(6, last + 1):
                v = ws.Cells(r, 1).Value
                if isinstance(v, str):
                    s = v.strip()
                    if s == "地图" or s.startswith("地图 "):
                        title_row = r
                        break
        except Exception:
            pass
        self._map_title_row = title_row
        self._map_header_row = title_row + 1
        self._map_row0 = title_row + 2
        header_row = self._map_header_row
        row0 = self._map_row0
        title = f"地图 {self.axis.map_code or ''}".strip()
        self._cell(ws, title_row, 1, title, bold=True, size=13)
        if not data:
            self._cell(ws, row0, 1, "没有地图数据（检查“关卡代号”）",
                       size=11)
            return
        height = int(data.get("height", 0))
        width = int(data.get("width", 0))
        tiles = data.get("tiles") or []
        if height <= 0 or width <= 0 or not tiles:
            self._cell(ws, row0, 1, "地图数据为空", size=11)
            return
        # 列号 1..width，行字母（A 在下）
        for j in range(width):
            self._cell(ws, header_row, MAP_COL0 + j, j + 1,
                       bold=True, size=10, border=True)
            ws.Columns(MAP_COL0 + j).ColumnWidth = 5.5
        for i in range(height):
            letter = chr(ord("A") + height - 1 - i)
            self._cell(ws, row0 + i, 1, letter, bold=True, size=10,
                       border=True)
            for j in range(width):
                try:
                    tile = tiles[i][j]
                    key = str(tile.get("tileKey", ""))
                except Exception:
                    key = ""
                fill = TILE_COLORS.get(key, DEFAULT_TILE)
                self._terrain[(row0 + i, MAP_COL0 + j)] = fill
                self._cell(ws, row0 + i, MAP_COL0 + j, "", fill=fill,
                           border=True, border_color="808080")
        for i in range(height):
            ws.Rows(row0 + i).RowHeight = 22
        # 图例
        legend_row = row0 + height + 1
        self._cell(ws, legend_row, 1, "图例", bold=True)
        items = [("tile_floor", "地面"), ("tile_road", "通道"),
                 ("tile_wall", "墙"), ("tile_forbidden", "不可部署"),
                 ("tile_start", "起点"), ("tile_end", "终点")]
        for k, (key, label) in enumerate(items):
            r = legend_row + 1 + k // 3
            c = 1 + (k % 3) * 2
            self._cell(ws, r, c, "", fill=TILE_COLORS.get(key, DEFAULT_TILE),
                       border=True)
            self._cell(ws, r, c + 1, label, size=9)
        self._cell(ws, legend_row, 8, "蓝=已部署  黄=当前目标", size=9)

    def _tile_cell(self, tile_pos, height: int):
        """convert_position 的 (number-1, letter-based) -> Excel 单元格。"""
        row_num, col_letter = tile_pos
        i = int(col_letter)
        j = int(row_num)
        label = chr(ord("A") + height - 1 - i) + str(j + 1)
        return self._map_row0 + i, MAP_COL0 + j, label

    def _flush_map(self, force: bool = False) -> bool:
        ws = self._ws_map
        data = self.axis.map_data
        if ws is None or not data:
            return False
        height = int(data.get("height", 0))
        if height <= 0:
            return False
        with self._lock:
            deployed = dict(self._deployed)
            action = self._current_action
            finished = self._finished
        # 用签名判断地图内容是否真的变了。变化才碰 Excel，避免每轮
        # 全量重画造成闪烁；同时也能捕捉到漏置 _map_dirty 的状态变化。
        marks = [
            ("D", str(oper), str(tile_pos), str(direction), str(name))
            for oper, (tile_pos, direction, name) in deployed.items()
        ]
        if action is not None and not finished:
            marks.append((
                "T", str(action.tile_pos), str(action.direction),
                str(action.alias or action.oper or action.kind),
            ))
        signature = tuple(sorted(marks))
        if not force and signature == self._map_signature:
            return False
        # 先清掉上一次的标记
        for (r, c) in self._marked:
            fill = self._terrain.get((r, c), DEFAULT_TILE)
            self._cell(ws, r, c, "", fill=fill, border=True,
                       border_color="808080")
        self._marked.clear()

        def mark(tile_pos, text, fill, border_color):
            if tile_pos is None:
                return
            try:
                r, c, _ = self._tile_cell(tile_pos, height)
            except Exception:
                return
            self._cell(ws, r, c, text, fill=fill, bold=True, size=9,
                       border=True, border_color=border_color)
            self._marked.add((r, c))

        for oper, (tile_pos, direction, name) in deployed.items():
            short = (name or oper)[:4]
            mark(tile_pos, short + ARROWS.get(direction or "", ""),
                 DEPLOYED_COLOR, "2E75B6")
        if action is not None and not finished:
            short = (action.alias or action.oper or action.kind)[:4]
            mark(action.tile_pos, short + ARROWS.get(action.direction or "", ""),
                 TARGET_COLOR, "BF8F00")
        self._map_signature = signature
        return True

    # ---------- JSON 兜底 ----------
    def _write_json(self) -> None:
        try:
            with self._lock:
                data = {
                    "axis": self.axis.path,
                    "report": self.report_path,
                    "status": self._status,
                    "error": self._error,
                    "frame": self._frame_text,
                    "time": self._time_text,
                    "cost_frame": self._cost_frame_text,
                    "rate": self._rate_text,
                    "results": {
                        str(k): {
                            "target": v.get("target"),
                            "actual": v.get("actual"),
                            "diff": v.get("diff"),
                            "note": v.get("note"),
                        } for k, v in self._results.items()
                    },
                }
            with open(self._json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
