"""WAA 交互式控制台。

启动后不要求先给 --axis；用 load 选轴并打开 Excel 面板，
write 选 F5 写入文件，run 读轴后自动进图并开始执行。
F5/F6 全局热键仍然可用。
"""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import shutil
import subprocess
import sys
import threading
import time
import traceback

import cv2
import win32gui

if getattr(sys, "frozen", False):
    ROOT = os.path.dirname(os.path.abspath(sys.executable))
else:
    ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from waa import VERSION_CHANNEL, __version__  # noqa: E402
from waa.core.axis import read_axis  # noqa: E402
from waa.core.capture import client_rect  # noqa: E402
from waa.core.capture_stream import CaptureStream  # noqa: E402
from waa.core.cost_model import rate_to_frames  # noqa: E402
from waa.core.excel_report import (  # noqa: E402
    ExcelReporter, _cost_frame_text, _frame_time_text,
)
from waa.core.executor import Executor, FatalReloadError, StopRequested  # noqa: E402
from waa.core.frame_counter import normalize_frame, start_screen_score  # noqa: E402
from waa.core.mark_writer import MarkWriter  # noqa: E402
from waa.core.ocr_name import ocr_text_boxes  # noqa: E402
from waa.paths import app_root, asset_path, bundle_root  # noqa: E402
from waa.core.resource_update import ResourceUpdater  # noqa: E402
from waa.core.run_log import RunLogger  # noqa: E402
from waa.core.send_input import click as mouse_click  # noqa: E402
from waa.core.touch_input import TouchDevice  # noqa: E402
from waa.core.vision import identify_cards, preload_card_templates  # noqa: E402
from waa.core.waa_counter import (  # noqa: E402
    RESET_AFTER_OUT_OF_BATTLE_S, WAACounter, FrameCounterThread,
)
from tools.run_axis import (  # noqa: E402
    OUT, _focus_window, _resolve_negative, _resolve_rate_config,
    _setting_bool, start_pause_hotkey,
)

CONFIG_PATH = os.path.join(OUT, "console_config.json")
# 随包只读资源（打包后在 _MEIPASS 里，开发时就是项目根）
ASSET_ROOT = bundle_root()
STANDARD_TEMPLATE = os.path.join(
    ASSET_ROOT, "waa", "templates", "WAA_标准模板.xlsx")
DEFAULT_ALPHA = 0.55
MIN_ALPHA = 0.10
MAX_ALPHA = 1.0
SHOWFRAME_INTERVAL = 0.01
PREVIEW_RATE = 100.0
PREVIEW_RATES = (50.0, 100.0, 110.0)
# 编队界面右下角找“开始行动”。不同版本按钮样式会变，参考 MAA 用 OCR，
# 不再依赖旧的固定模板/固定坐标。ROI 是 1920×1080 下的像素。
START_TEXTS = ("开始行动", "开始作战", "开始推演", "开始突袭",
               "开始试炼", "开始测试", "开始")
START_BUTTON_ROI = (1000, 550, 920, 530)
START_BUTTON_CHARS = set("开始行动作战推演突袭试炼测试")


def _same_path(a: str, b: str) -> bool:
    return (os.path.normcase(os.path.abspath(a))
            == os.path.normcase(os.path.abspath(b)))


def _parse_alpha(text: str) -> float | None:
    """把 showframe 的数字转成 0.10~1.0 的透明度。

    55 / 55% -> 0.55；0.55 -> 0.55；100 -> 1.0。
    """
    raw = text.strip().replace("%", "")
    try:
        value = float(raw)
    except ValueError:
        print("透明度请输入 10~100 的数字，例如 showframe 55")
        return None
    # 带小数点且 <=1 的写法按比例；其余按百分比。
    alpha = value if (0 < value <= 1 and "." in raw) else value / 100.0
    if not (MIN_ALPHA <= alpha <= MAX_ALPHA):
        print("透明度范围 10~100，例如 showframe 55")
        return None
    return alpha


def _make_preview_counter(start_frame: int = 0,
                          require_hud: bool = False) -> WAACounter:
    """没 load 轴时给 showframe 用的独立计数（默认 100%，自动识别）。"""
    counter = WAACounter(PREVIEW_RATE, rate_candidates=list(PREVIEW_RATES),
                         require_hud=require_hud)
    counter.arm(max(0, int(start_frame)))
    return counter


def _pick_start_point(lines) -> tuple[tuple[int, int], str] | None:
    """从 OCR 行里选“开始行动”按钮中心（1920×1080 坐标）。

    优先取“开始/行动”等字的外接框，避开按钮旁边的图标、加号和英文小字；
    版本换皮后只要文字还在，就能点中。
    """
    best = None
    for line in lines:
        text = "".join((line.get("text") or "").split())
        if not any(key in text for key in START_TEXTS):
            continue
        boxes = []
        for word in line.get("words") or []:
            wt = "".join((word.get("text") or "").split())
            if any(ch in START_BUTTON_CHARS for ch in wt):
                boxes.append(word)
        if not boxes:
            boxes = [line]
        x0 = min(int(b.get("x", 0)) for b in boxes)
        y0 = min(int(b.get("y", 0)) for b in boxes)
        x1 = max(int(b.get("x", 0)) + int(b.get("w", 0))
                 for b in boxes)
        y1 = max(int(b.get("y", 0)) + int(b.get("h", 0))
                 for b in boxes)
        if x1 <= x0 or y1 <= y0:
            continue
        area = (x1 - x0) * (y1 - y0)
        if best is None or area > best[0]:
            best = (area, ((x0 + x1) // 2, (y0 + y1) // 2),
                    line.get("text") or "")
    return (best[1], best[2]) if best is not None else None


def _load_config() -> dict:
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {}


def _save_config(data: dict) -> None:
    try:
        os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


class HotkeyProxy:
    """F5/F6 热键只注册一次，load 换轴后转发到新执行器。"""

    def __init__(self):
        self.executor: Executor | None = None

    def mark_frame(self):
        if self.executor is None:
            print("F5：还没有 load 轴文件")
            return None
        return self.executor.mark_frame()

    def toggle_manual_pause(self):
        if self.executor is None:
            print("F6：还没有 load 轴文件")
            return
        return self.executor.toggle_manual_pause()


class ShowFrameWindow:
    """独立进程的三行半透明帧窗口。

    Tk 在主进程的非主线程里 destroy 会触发 Tcl_AsyncDelete；所以窗口
    放在独立进程，主进程只写状态文件。
    """

    def __init__(self, snapshot, alpha: float = DEFAULT_ALPHA):
        self.snapshot = snapshot
        self.alpha = self._clamp_alpha(alpha)
        self.proc = None
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self._state_lock = threading.Lock()
        self._last_payload = None
        self.state_path = os.path.join(OUT, "showframe_state.json")
        self.script = os.path.join(ROOT, "tools", "frame_window.py")

    @staticmethod
    def _clamp_alpha(value: float) -> float:
        try:
            value = float(value)
        except (TypeError, ValueError):
            value = DEFAULT_ALPHA
        return max(MIN_ALPHA, min(MAX_ALPHA, value))

    def set_alpha(self, value: float) -> float:
        self.alpha = self._clamp_alpha(value)
        self._write_state()
        return self.alpha

    def show(self) -> None:
        if self.proc is not None and self.proc.poll() is None:
            return
        if self.proc is not None:
            # 用户右键关掉了窗口：先把旧写状态线程停掉。
            self.hide()
        os.makedirs(OUT, exist_ok=True)
        self._stop.clear()
        self._last_payload = None
        self._write_state()
        self._thread = threading.Thread(target=self._write_loop, daemon=True)
        self._thread.start()
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        if getattr(sys, "frozen", False):
            # 打包后没有 pythonw：帧窗是随包发布的第二个 exe。
            cmd = [os.path.join(ROOT, "WAA_frame.exe"),
                   self.state_path, str(os.getpid())]
        else:
            python = sys.executable
            pythonw = os.path.join(os.path.dirname(python), "pythonw.exe")
            if not os.path.exists(pythonw):
                pythonw = python
            cmd = [pythonw, self.script, self.state_path, str(os.getpid())]
        try:
            self.proc = subprocess.Popen(cmd, creationflags=flags)
        except Exception as exc:
            print(f"showframe 启动失败：{exc}")
            self._stop.set()

    def _write_state(self) -> None:
        try:
            with self._state_lock:
                frame, time_text, cost = self.snapshot()
                payload = {"frame": frame, "time": time_text,
                           "cost": cost, "alpha": round(self.alpha, 4)}
                if payload == self._last_payload:
                    return
                os.makedirs(os.path.dirname(self.state_path), exist_ok=True)
                tmp = self.state_path + ".tmp"
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(payload, f, ensure_ascii=False)
                os.replace(tmp, self.state_path)
                self._last_payload = payload
        except Exception:
            pass

    def _write_loop(self) -> None:
        while not self._stop.is_set():
            self._write_state()
            self._stop.wait(SHOWFRAME_INTERVAL)

    def hide(self) -> None:
        self._stop.set()
        if self._thread is not None and self._thread.is_alive():
            self._thread.join(timeout=1.0)
        self._thread = None
        if self.proc is not None:
            try:
                self.proc.terminate()
                self.proc.wait(timeout=2.0)
            except Exception:
                pass
            self.proc = None

    def stop(self) -> None:
        self.hide()


class _PromptGuard:
    """等输入时把后台线程的输出攒起来，避免糊在「waa> 」那一行上。

    现象：input() 打出提示符后，vision/热键等后台线程正好 print，
    于是变成 `waa> [vision] ...`，接着用户敲的命令就落在没有提示符的下一行。

    做法：主线程阻塞在 input() 期间 hold 住，非主线程的写入先进缓冲；
    用户按回车后 release，把缓冲原样吐出来（此时提示符那行已经结束了）。
    主线程自己的写入（提示符本身）永远直通。
    """

    def __init__(self, stream):
        self._stream = stream
        self._lock = threading.Lock()
        self._held = False
        self._pending: list[str] = []

    def hold(self) -> None:
        with self._lock:
            self._held = True

    def release(self) -> None:
        with self._lock:
            self._held = False
            pending, self._pending = self._pending, []
        if not pending:
            return
        try:
            for chunk in pending:
                self._stream.write(chunk)
            self._stream.flush()
        except Exception:
            pass

    def write(self, data):
        if not data:
            return 0
        if (threading.current_thread() is not threading.main_thread()):
            with self._lock:
                if self._held:
                    self._pending.append(data)
                    return len(data)
        self._stream.write(data)
        return len(data)

    def flush(self):
        try:
            self._stream.flush()
        except Exception:
            pass

    def __getattr__(self, name):
        return getattr(self._stream, name)


class ConsoleSession:
    def __init__(self, args):
        self.args = args
        self.hwnd = win32gui.FindWindow("UnityWndClass", "明日方舟")
        if not self.hwnd:
            raise RuntimeError("找不到游戏窗口")
        self.stream = CaptureStream(self.hwnd)
        self.stream.start()
        self.touch = TouchDevice()
        self.axis = None
        self.axis_path = None
        self.counter = None
        self.counter_thread = None
        self.reporter = None
        self.executor = None
        self.run_thread = None
        self.running = False
        # 轴跑完后 run 线程不退出，接着服务 F6（复用同一个线程）。
        self.f6_service_active = False
        self.f6_stop = threading.Event()
        # 退出关卡自动 stop：run 会话期间盯着计数器的“离开战斗”状态，
        # HUD 消失超过 3 秒就自动结束本次 run（等同输入 stop）。
        self._watcher_stop = threading.Event()
        self._battle_watcher = None
        self._battle_seen = False
        self._auto_stop_reason = ""
        self.reading = False
        self.write_path = None
        self.write_writer = None
        self.preview_ocr = None
        self.preview_offset = int(self.args.start_offset)
        self._auto_enter_failure = ""
        self.run_logger = None
        self.log_enabled = False   # 日志默认关闭，输入 log on 才记录
        self._verbose = False      # 调试输出默认关闭，输入 verbose on 才打印细节
        self._hud_tpl = None
        self._closed = False
        self.config = _load_config()
        self.default_dir = self.config.get("default_dir") or os.path.join(
            ROOT, "axes")
        self.hotkeys = HotkeyProxy()
        start_pause_hotkey(self.hotkeys)
        self.updater = ResourceUpdater()
        if not getattr(self.args, "no_update", False):
            self.updater.start_background(self._report_auto_update)
        self.showframe = ShowFrameWindow(
            self.snapshot_frame,
            alpha=self.config.get("showframe_alpha", DEFAULT_ALPHA))
        self.auto_showframe = bool(self.config.get("showframe_auto", True))
        # 不 load 也能看帧：启动时就跑一个预览计数器。
        start_frame = (int(self.args.start_frame)
                       if self.args.start_frame is not None else 0)
        self._start_preview_counter(start_frame)
        # 头像库约 10 秒；启动时后台预热，不要等第一次部署才加载。
        self._card_preload_thread = threading.Thread(
            target=self._preload_cards, daemon=True)
        self._card_preload_thread.start()

    @staticmethod
    def _preload_cards() -> None:
        preload_card_templates(log=lambda msg: print(f"[vision] {msg}"))

    # ---------- 状态 ----------
    def snapshot_frame(self):
        ex, c = self.executor, self.counter
        if ex is not None:
            cur = ex.current_frame()
        elif c is not None:
            cur = c.frame + self.preview_offset
        else:
            cur = None
        if cur is None:
            return "—", "—", "—"
        return (str(cur), _frame_time_text(cur),
                _cost_frame_text(c.last) if c is not None else "—")

    # ---------- 运行时 ----------
    @staticmethod
    def _report_auto_update(message: str) -> None:
        print(f"[update] {message}")

    def update_command(self) -> None:
        if self.running:
            print("轴正在运行，先 stop 再 update")
            return
        print("正在从 GitHub 检查本地游戏数据更新…")
        result = self.updater.update(force=True, quiet=True)
        print(f"[update] {result.get('message', '')}")
        if result.get("status") == "updated":
            print("数据已更新；如果已经 load 过轴，请重新 load 以使用新地图/头像。")

    def _start_preview_counter(self, start_frame: int = 0) -> None:
        """没有 load 轴时，用独立计数器让 showframe 直接显示帧号。"""
        if self.counter is not None:
            return
        self.counter = _make_preview_counter(start_frame, require_hud=True)
        self.counter_thread = FrameCounterThread(self.stream, self.counter)
        self.counter_thread.start()

    def _current_frame_state(self) -> tuple[int, int]:
        """当前原始帧和 offset；load 时用它们接着跑，而不是回到 0。"""
        raw = int(self.counter.frame) if self.counter is not None else 0
        if self.executor is not None:
            offset = int(self.executor.start_offset)
        else:
            offset = int(self.preview_offset)
        return raw, offset

    def _teardown_runtime(self) -> None:
        self._stop_f6_service()
        if self.run_thread is not None and self.run_thread.is_alive():
            if self.executor is not None:
                self.executor.request_stop()
            self.run_thread.join(timeout=8.0)
        if self.reporter is not None:
            try:
                self.reporter.stop()
            except Exception:
                pass
        if self.counter_thread is not None:
            try:
                self.counter_thread.stop()
            except Exception:
                pass
        if self.counter is not None:
            try:
                self.counter.close_log()
            except Exception:
                pass
        self.executor = None
        self.reporter = None
        self.counter = None
        self.counter_thread = None
        self.run_thread = None
        self.running = False
        self.hotkeys.executor = None

    def _attach_write_writer(self) -> None:
        if self.executor is None:
            return
        if self.write_path is None:
            if self.write_writer is not None:
                self.write_writer.close()
                self.write_writer = None
            self.executor.mark_writer = None
            self.executor.mark_writer_error = ""
            return
        if (self.axis_path is not None
                and _same_path(self.write_path, self.axis_path)):
            # 同一个文件：F5 走主轴的 ExcelReporter，避免两个写入者。
            if self.write_writer is not None:
                self.write_writer.close()
                self.write_writer = None
            self.executor.mark_writer = None
            self.executor.mark_writer_error = ""
            return
        if self.write_writer is None:
            try:
                self.write_writer = MarkWriter(self.write_path)
            except Exception as exc:
                print(f"write 文件打开失败：{exc}")
                self.executor.mark_writer = None
                self.executor.mark_writer_error = (
                    f"write 文件打不开：{exc}")
                return
        self.executor.mark_writer = self.write_writer
        self.executor.mark_writer_error = ""

    def load_axis(self, path: str) -> bool:
        if self.running:
            print("轴正在运行，先 stop 再 load")
            return False
        path = os.path.abspath(path)
        if not os.path.exists(path):
            print(f"找不到轴文件：{path}")
            return False
        self.reading = True
        if self.executor is not None:
            self.executor.mark_suspended = True
        resume_raw, resume_offset = self._current_frame_state()
        try:
            axis = read_axis(path)
            rate, neg_speed, neg_declared, rate_src, declared = (
                _resolve_rate_config(axis, self.args.cost_rate))
            if rate <= 0:
                print("回费速度为 0，无法计帧")
                return False
            neg_override, neg_src = _resolve_negative(
                axis, self.args.negative)
            resume_raw, resume_offset = self._current_frame_state()
            # 换轴时关闭上一个由本程序打开的 Excel，避免留下多个
            # 隐藏 Excel/COM apartment；用户自己打开的 Excel 不会被关。
            if self.reporter is not None:
                self.reporter.quit_on_stop = True
            self._teardown_runtime()
            self.preview_offset = resume_offset
            self.axis = axis
            self.axis_path = path
            candidates = (declared if (declared
                                       and self.args.cost_rate is None)
                          else [rate])
            self.counter = WAACounter(
                rate, neg_override=neg_override, neg_speed=neg_speed,
                neg_declared=neg_declared, rate_candidates=candidates,
                require_hud=True)
            self.counter.csv_path = os.path.join(OUT, "waa_run.csv")
            # 从控制台已经记到的帧继续，不回到 0。
            self.counter.arm(resume_raw)
            self.counter_thread = FrameCounterThread(
                self.stream, self.counter)
            self.counter_thread.start()
            template = (STANDARD_TEMPLATE if not self.args.copy else
                        os.path.join(ROOT, "axes", "WAA_sample_1-7.xlsx"))
            self.reporter = ExcelReporter(
                axis,
                {"base": rate, "declared": declared,
                 "neg_speed": neg_speed},
                os.path.join(OUT, "run_reports"),
                enabled=not self.args.no_excel,
                template_path=template,
                in_place=not self.args.copy,
                backup=False,
            )
            self.executor = Executor(
                self.hwnd, axis, self.stream, self.touch, self.counter,
                start_offset=resume_offset)
            self.executor.select_pause_mode = self.args.pause_mode
            self.executor.verbose = self._verbose
            if self.args.speed != "auto":
                self.executor.speed_override = int(self.args.speed)
            self.executor.step_enabled = not self.args.no_step
            self.executor.pause_method = self.args.pause_method
            self.executor.end_pause = _setting_bool(
                axis.settings.get("结束后暂停"), False)
            self.reporter.bind_runtime(self.counter, self.executor)
            self.executor.reporter = self.reporter
            self.hotkeys.executor = self.executor
            self._attach_write_writer()
            # load 后立即把 Excel 面板/地图画出来；run 会复用这个面板。
            if not self.args.no_excel:
                print("打开 Excel 面板…")
                if not self.reporter.start(wait_ready=True):
                    print("Excel 面板没打开；轴已加载，可关闭占用文件的 Excel 后 run 重试")
                else:
                    # load 就把顺序定下来：用户看到的就是执行顺序。
                    self.executor.sort_workbook()
            print(f"已加载：{path}")
            print(f"地图={axis.map_code} 动作={len(axis.actions)} "
                  f"回费={rate:g}% 候选={'/'.join(f'{x:g}' for x in candidates)}")
            print(f"局内负费={neg_src}")
            return True
        except Exception as exc:
            print(f"[错误] load 失败：{exc}")
            if self.counter is None:
                self.preview_offset = resume_offset
                self._start_preview_counter(resume_raw)
            return False
        finally:
            self.reading = False
            if self.executor is not None:
                self.executor.mark_suspended = False

    def _get_hud_template(self):
        if self._hud_tpl is None:
            self._hud_tpl = cv2.imread(
                os.path.join(ASSET_ROOT, "waa", "resources",
                             "BattleOfficiallyBegin.png"),
                cv2.IMREAD_GRAYSCALE)
        return self._hud_tpl

    def _start_screen_score(self, frame=None) -> float:
        if frame is None:
            frame = self.stream.latest()
        if frame is None:
            return -1.0
        gray = cv2.cvtColor(normalize_frame(frame), cv2.COLOR_BGR2GRAY)
        return start_screen_score(gray)

    def _find_start_button(self, frame=None):
        """返回 (score, 点击点, 方法)。

        先用 OCR 找“开始行动/开始”等文字并直接点文字中心；OCR 不可用时
        退回旧模板，模板命中但没坐标时由调用方用旧的固定比例坐标。
        点击点是 1920×1080 标准坐标。
        """
        if frame is None:
            frame = self.stream.latest()
        if frame is None:
            return -1.0, None, "none"
        norm = normalize_frame(frame)
        try:
            lines = ocr_text_boxes(norm, START_BUTTON_ROI)
            hit = _pick_start_point(lines)
            if hit is not None:
                point, text = hit
                return 1.0, point, f"ocr:{text}"
        except Exception as exc:
            print(f"开始按钮 OCR 失败（改用模板）：{exc!r}")
        return self._start_screen_score(norm), None, "template"

    def _hud_score(self, frame=None) -> float:
        if frame is None:
            frame = self.stream.latest()
        tpl = self._get_hud_template()
        if frame is None or tpl is None:
            return -1.0
        gray = cv2.resize(cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY),
                          (1280, 720))
        roi = gray[20:85, 1165:1240]
        if tpl.shape[0] > roi.shape[0] or tpl.shape[1] > roi.shape[1]:
            return -1.0
        return float(cv2.matchTemplate(
            roi, tpl, cv2.TM_CCOEFF_NORMED).max())

    def _wait_armed(self, timeout: float) -> bool:
        deadline = time.time() + timeout
        while time.time() < deadline and not self.counter.armed:
            time.sleep(0.05)
        return bool(self.counter.armed)

    def _in_battle(self) -> bool:
        """关卡内以 HUD 信号为准。

        打完一局回到编队界面后，counter.armed 可能还是 True；只看 armed
        会误以为还在战斗，run 就不会点“开始行动”。
        """
        return bool(getattr(self.counter, "in_battle", False))

    def _wait_battle(self, timeout: float) -> bool:
        deadline = time.time() + timeout
        while time.time() < deadline and not self._in_battle():
            time.sleep(0.05)
        return self._in_battle()

    def _auto_enter_stage(self) -> bool:
        """在编队界面点“开始行动”，等 HUD/费用条后返回 True。"""
        self._auto_enter_failure = ""
        if self.executor is None or self.reporter is None:
            self._auto_enter_failure = "运行组件未初始化"
            return False
        frame = None
        deadline = time.time() + 3.0
        while time.time() < deadline and frame is None:
            frame = self.stream.latest()
            if frame is None:
                time.sleep(0.02)
        if frame is None:
            self._auto_enter_failure = "3 秒内没有抓到游戏画面"
            print("3 秒内没有抓到游戏画面，已取消自动进图")
            return False
        score, click_pt, method = self._find_start_button(frame)
        print(f"编队界面检测 score={score:.3f} 方法={method}")
        if score < 0.55:
            self._auto_enter_failure = (
                f"未识别到编队界面的“开始行动”（score={score:.3f}）")
            if frame is not None:
                cv2.imwrite(os.path.join(OUT, "run_not_squad.png"), frame)
            return False
        if click_pt is not None:
            print(f"OCR 定位“开始行动”：{click_pt}")

        # 必须在点“开始行动”前打开 Excel，否则鼠标点击会点到 Excel 上。
        if (not self.args.no_excel
                and not getattr(self.reporter, "started", False)):
            print("开始作战前打开 Excel 面板…")
            if not self.reporter.start(wait_ready=True):
                detail = getattr(self.reporter, "start_error", "") or "未知原因"
                self._auto_enter_failure = (
                    f"Excel 面板没打开，已取消自动进图：{detail}")
                print(self._auto_enter_failure)
                return False

        max_attempts = 3
        per_wait = 10.0
        detected = False
        for attempt in range(1, max_attempts + 1):
            if attempt > 1:
                time.sleep(0.5)
                sc, pt, how = self._find_start_button(self.stream.latest())
                print(f"编队界面重试检测 score={sc:.3f} 方法={how}")
                if sc < 0.55:
                    print("当前不在编队界面（可能在加载），不再重复点击")
                    break
                if pt is not None:
                    click_pt = pt
            # 前置游戏窗口。_focus_window 内部会验证“真的到前台”才返回 True，
            # 这里再给它最多 3 次机会：偶尔会有别的窗口（Excel / 控制台）
            # 在被抢回前台，多试两次比直接取消整轮 run 划算。
            focused = False
            for _try in range(3):
                if _focus_window(self.hwnd):
                    focused = True
                    break
                time.sleep(0.12)
            if not focused:
                fg = win32gui.GetForegroundWindow()
                try:
                    fg_title = win32gui.GetWindowText(fg) or "（无标题）"
                except Exception:
                    fg_title = "?"
                self._auto_enter_failure = (
                    f"游戏窗口没能切到前台（当前前台：{fg_title}）")
                print(self._auto_enter_failure)
                return False
            left, top, cw, ch = client_rect(self.hwnd)
            if click_pt is not None:
                # OCR 坐标是 1920×1080 标准坐标，按客户区缩放。
                sx = left + int(click_pt[0] / 1920.0 * cw)
                sy = top + int(click_pt[1] / 1080.0 * ch)
            else:
                sx = left + int(0.862 * cw)
                sy = top + int(0.700 * ch)
            print(f"foreground={_is_foreground(self.hwnd)} "
                  f"client={cw}x{ch} click=({sx},{sy})")
            mouse_click(sx, sy)
            print(f"第 {attempt}/{max_attempts} 次点击“开始行动” "
                  f"({sx},{sy})，等待战斗 HUD…")
            deadline = time.time() + per_wait
            while time.time() < deadline:
                hud = self._hud_score(self.stream.latest())
                # 旧 HUD 模板可能跟不上新版本；计数器自己的按钮/费用条
                # 识别同样能确认已经进图。
                if (hud >= 0.65 or getattr(self.counter, "in_battle", False)
                        or getattr(self.counter, "armed", False)):
                    detected = True
                    break
                time.sleep(0.02)
            if detected:
                break
            print(f"第 {attempt} 次点击后 {per_wait:.0f} 秒未检测到战斗 HUD")

        if not detected:
            print("最后等待战斗 HUD（不再点击）…")
            deadline = time.time() + 10.0
            while time.time() < deadline:
                hud = self._hud_score(self.stream.latest())
                if (hud >= 0.65 or getattr(self.counter, "in_battle", False)
                        or getattr(self.counter, "armed", False)):
                    detected = True
                    break
                time.sleep(0.02)
        if not detected:
            frame = self.stream.latest()
            if frame is not None:
                cv2.imwrite(os.path.join(OUT, "run_no_hud.png"), frame)
            self._auto_enter_failure = "点了开始行动，但没有检测到战斗 HUD"
            print(self._auto_enter_failure + "，已取消自动进图")
            return False

        # 新关卡从 0 开始；offset 保留。
        self.counter.arm(0)
        # 开局暂停照搬 AFA：等倍速按钮（speed_drawn）一画出就发一次 Space、
        # 不确认。倍速按钮比费用条（armed）早约 0.8 秒，所以这里发完 Space 后
        # armed 可能还没翻转，后面再等它出现用于确认“已进关”。
        try:
            if self.executor is not None:
                print("等倍速按钮出现后发 Space 暂停…")
                self.executor.pause_by_key()
        except Exception as exc:
            print(f"进关暂停失败（继续）：{exc!r}")
        # pause_by_key 发 Space 时 armed 可能还没出现；这里等它出现确认进关。
        if self.counter.armed and self.counter.last is not None:
            print("已进入战斗，费用条已出现")
            return True
        # 极少数情况：倍速按钮到费用条之间超过 2 秒，再给 2 秒确认是否进关。
        deadline = time.time() + 2.0
        while time.time() < deadline:
            if self.counter.armed and self.counter.last is not None:
                print("已进入战斗，费用条已出现")
                return True
            time.sleep(0.05)
        self._auto_enter_failure = "进关后费用条一直没出现"
        print(self._auto_enter_failure)
        return False

    def _start_execution(self, auto_enter: bool = False) -> None:
        if self.executor is None or self.axis is None:
            print("先 load 一个轴文件")
            return
        if self.running:
            print("上一次 run 还没结束：先输入 stop 结束它（日志会保存），再 run")
            return
        # 新的 run 接管暂停控制权，先关掉上一轮留下的 F6 守护。
        self._stop_f6_service()
        # 头像库必须在开跑前加载完：首次构建约 10 秒，而选人/部署的帧窗口
        # 只有几百毫秒，边跑边加载会直接毁掉那一段（表现为“所有入口都没进入
        # 子弹时间”）。缓存没有就在这里现场构建，不通过就干脆不开跑。
        if not preload_card_templates(log=lambda m: print(f"[vision] {m}"),
                                     force=True):
            print("头像库没加载成功，已取消 run")
            _bring_console_to_front()
            return
        print(f"run: frame={self.executor.current_frame()} "
              f"armed={self.counter.armed} "
              f"in_battle={getattr(self.counter, 'in_battle', None)}")
        # 编队界面里 armed 可能是上一局残留的 True，必须看 in_battle。
        self._auto_enter_failure = ""
        if not self._in_battle():
            if auto_enter:
                self._wait_battle(0.6)
                if not self._in_battle():
                    if not self._auto_enter_stage():
                        reason = self._auto_enter_failure or "自动进图失败"
                        print(f"未自动进图：{reason}；继续等待进入战斗…")
                        self._wait_battle(4.4)
            else:
                self._wait_armed(5.0)
        if not self._in_battle():
            if self._auto_enter_failure:
                print(f"run 未开始：{self._auto_enter_failure}")
                # 自动化进图失败时游戏通常还在前台，控制台压在后台，
                # 用户只看到“什么也没发生”。把原因摆到眼前。
                _bring_console_to_front()
            else:
                print("还没进入战斗：请先进入关卡，再 run")
            return
        if not self.counter.armed:
            self._wait_armed(4.0)
        if not self.counter.armed:
            print("费用条还没出现：请等费用条显示后再 run")
            return
        if (not self.args.no_excel
                and not getattr(self.reporter, "started", False)):
            print("打开 Excel 面板…")
            if not self.reporter.start(wait_ready=True):
                print("Excel 面板没打开，已取消 run")
                _bring_console_to_front()
                return
        if not self.args.no_excel and self.executor is not None:
            # run 之前再排一遍：load 之后可能又补了行，或者在 Excel 里手工
            # 调过顺序。已经有序时 sort_workbook 只读一遍就返回，几乎不耗时。
            self.executor.sort_workbook()
        if not _focus_window(self.hwnd):
            print("WARNING: 游戏窗口没到前台")
        try:
            # 开局暂停只有一处：进关时 _auto_enter_stage 已经用 Space 停住了
            # （照搬 AFA，发一次不确认）。这里**不再点第二下**——视觉样本比游戏
            # 晚 40~100ms，刚停住时读到的还是“运行中”，一点就把游戏点开
            # （2026-11-09 06:16：帧 2 停住 → 被点开 → 帧 7 才停）。Space 万一没
            # 生效，交给第一个 pause_at 兜底真正暂停。
            default_2x = _setting_bool(
                self.axis.settings.get("默认2倍速"), False)
            if self.args.speed == "auto" and default_2x:
                if not self.executor.ensure_speed(2):
                    print("WARNING: 默认 2 倍速切换失败")
            if _setting_bool(self.axis.settings.get("开始前暂停"), False):
                self.executor.manual_pause.set()
                print("开始前暂停：游戏保持暂停，按 F6 或 resume 恢复")
            # 没勾“开始前暂停”也不在这里 resume：开局暂停成功就一直停着，
            # 由轴自己的 pause_at 决定什么时候开跑（它需要跑帧时才恢复），
            # 这样不会因为“先恢复几十毫秒再重新暂停”白丢帧。
        except Exception as exc:
            print(f"run 初始化失败：{exc}")
            return
        self.running = True
        # 退出关卡自动 stop：守护线程跟着整个 run 会话走（含轴跑完后的
        # F6 待命阶段），_end_run_session 里统一停。
        self._start_battle_watcher()
        total = len(self.axis.actions) or 1
        self.run_thread = threading.Thread(
            target=self._run_worker, daemon=True)
        self.run_thread.start()
        print(f"开始执行：轴共 {total} 行")

    def _run_worker(self) -> None:
        stopped = False
        try:
            total = len(self.axis.actions) or 1
            self.executor.run(max_actions=total,
                              start_action=self.args.start_action)
            if self.executor.stop_requested.is_set():
                stopped = True
            self.reporter.finish(not stopped)
        except FatalReloadError as exc:
            self.reporter.set_error(str(exc))
            self.reporter.finish(False)
            print(f"[错误] 轴已停止：{exc}")
            _bring_console_to_front()
        except Exception as exc:
            self.reporter.set_error(f"轴运行失败：{exc}")
            self.reporter.finish(False)
            print(f"[错误] 轴运行失败：{type(exc).__name__}: {exc}")
            # 完整堆栈只写进日志文件（若已开启），不刷屏。
            if self.run_logger is not None:
                print(traceback.format_exc(), file=sys.stderr)
            # 报错停止时控制台可能在后台（跑图时游戏在前台），
            # 不前置的话用户只看到游戏停住，看不到错误。
            _bring_console_to_front()
        finally:
            # run 是「会话」，不是「一次轴执行」：轴跑完只表示到了判断
            # “结束后暂停”的时间点，后面还要靠 F6 暂停/前置 Excel/重载继续。
            # 所以这里**不关日志**，只有 stop / quit 才收尾。
            if self.executor.stop_requested.is_set():
                print("轴已按 stop 退出：结束本次 run")
                self._end_run_session("stopped")
            else:
                print("轴执行完毕，本次 run 仍在进行")
                print("  F6：暂停 + 前置 Excel（补行）→ 再按一次：重载轴并接着跑")
                print("  stop：结束本次 run（日志保存）")
        if self.running and not self.executor.stop_requested.is_set():
            # 不新开线程：让这个 run 线程接着消费 F6 请求。
            self._f6_service_loop()

    def _f6_service_loop(self) -> None:
        """轴跑完后**复用 run 线程**继续服务 F6，直到 stop / quit。

        不走新线程：F6 的老路径本来就是「热键线程只登记事件、run 线程
        在安全点消费」；轴跑完 run 线程一退，事件就没人管了。这里让
        run 线程不退出，直接接着消费，语义和轴运行时完全一致：
        暂停 → 打开/前置 Excel → 再按 F6 保存并重载轴 → 恢复游戏。

        注意：这次 run 的日志此时**仍然开着**，F6 的每次操作都会写进去。
        """
        executor = self.executor
        if executor is None:
            return
        self.f6_service_active = True
        print("F6：暂停 + 前置 Excel（补行）→ 再按一次：重载轴并接着跑；"
              "stop：结束本次 run")
        last = executor.manual_pause.is_set()
        try:
            while not self.f6_stop.wait(0.03):
                now = executor.manual_pause.is_set()
                if now == last:
                    continue
                last = now
                if not now:
                    continue          # 只在“按下”的那一刻动作
                try:
                    # 第一下 F6：停住（已暂停就不动）+ 打开/前置 Excel，
                    # 让你在里面补行。_manual_hold 会一直等到第二次 F6。
                    executor._manual_hold()
                    if executor.manual_pause.is_set():
                        continue
                    # 第二下 F6：用刚保存的轴接着跑（新写进去的行会执行）
                    self._resume_and_rerun()
                    return
                except StopRequested:
                    break
                except Exception as exc:
                    # 重载轴失败不该把 F6 一起废掉：暂停/恢复本身已经生效。
                    print(f"  F6 处理失败（{type(exc).__name__}）：{exc}")
                    time.sleep(0.2)
        finally:
            self.f6_service_active = False

    def _resume_and_rerun(self) -> None:
        """F6 恢复：重读轴 → 从当前帧接着跑。

        已过帧的行会自动跳过，所以只有 F5 新写进去的那几行（帧号 ≥ 当前帧）
        会被执行——这就是「同一局里跑新写的行」。

        **不要在这里提前 resume**：F5 打轴记的是打轴那一刻的当前帧（例如
        「结束后暂停」时帧冻结在 1111）。若这里先 resume，游戏会立刻空跑、
        帧号涨过 1111，新写的行就全被「已过帧」跳过（2026-11-09 08:00 轮：
        新加 1111/1112 两行全没执行）。resume 交给 run 线程自己——它的第一个
        动作会通过 pause_at 判断：已暂停且就在目标帧则直接执行，离目标还远
        才 resume 跑过去。
        """
        path = self.axis_path
        if not path:
            self._stop_f6_service()
            return
        try:
            axis = read_axis(path)
        except Exception as exc:
            print(f"  F6：重读轴失败，保持暂停：{exc!r}")
            self._stop_f6_service()
            return
        self.axis = axis
        self.executor.axis = axis
        # reporter 内部也持有一份 axis 引用（load 时和 executor 指向同一个对象，
        # 但这里 read_axis 换成了新对象），不同步的话 reporter 的 axis.actions 还是
        # 旧长度，新加行的 index 会越界，序号/运行结果都写不上
        # （2026-11-09 08:23 轮：IndexError index=14，撤退斑点行没写序号和结果）。
        if self.reporter is not None:
            self.reporter.axis = axis
        self.f6_service_active = False      # 免得 _start_* 停到自己
        total = len(axis.actions) or 1
        print(f"F6：用刚保存的轴继续跑（共 {total} 行；已过帧的行自动跳过）")
        # 不再 start_run_log：本次 run 的日志一直开着，继续往同一个文件写。
        self.running = True
        self.run_thread = threading.Thread(
            target=self._run_worker, daemon=True)
        self.run_thread.start()

    def _stop_f6_service(self) -> None:
        """结束 F6 服务循环（只在这个循环真的在跑时才动作）。"""
        if not self.f6_service_active:
            return
        self.f6_stop.set()
        th = self.run_thread
        # 可能就是从服务循环自己里调过来的（F6 重读轴失败等），不能 join 自己
        if th is not None and th.is_alive() and th is not threading.current_thread():
            th.join(timeout=2.0)
        self.f6_stop.clear()

    # ---------- 退出关卡自动 stop ----------

    def _battle_exit_should_stop(self) -> bool:
        """检测“已退出关卡”：先见过战斗，之后 HUD 消失够久。

        判据和帧号归零是同一个信号源：计数器确认离开战斗
        （HUD/费用条连续多帧不见）持续 ≥ RESET_AFTER_OUT_OF_BATTLE_S。
        部署视角、UI 过渡造成的短暂丢失（1 秒左右）不会触发。

        必须先“真的见过战斗”才开始判退出：run 之后还没进图、加载画面
        超过 3 秒，都不能算退出关卡。**不能用 counter.armed**——编队界面
        里它可能是上一局残留的 True，得看实时的 in_battle。
        """
        counter = self.counter
        if counter is None:
            return False
        if not getattr(self, "_battle_seen", False):
            if getattr(counter, "in_battle", False):
                self._battle_seen = True
            return False
        dur = counter.out_of_battle_s()
        return dur >= RESET_AFTER_OUT_OF_BATTLE_S

    def _battle_watcher_loop(self) -> None:
        """run 会话期间的后台守护：退出关卡 -> 自动 stop 本次 run。"""
        ev = self._watcher_stop
        while not ev.wait(0.5):
            try:
                if not self._battle_exit_should_stop():
                    continue
                dur = self.counter.out_of_battle_s()
                print(f"\n检测到已退出关卡（HUD 消失 {dur:.1f}s）："
                      f"自动结束本次 run")
                self._auto_stop_reason = "out_of_battle"
                self.stop()
                return
            except Exception as exc:
                # 守护线程不能因为一次读数异常把会话搞挂
                print(f"  [watcher] 检测异常（忽略）：{exc!r}")

    def _start_battle_watcher(self) -> None:
        """run 会话开始时启动守护线程（已在跑就只重置状态）。"""
        self._battle_seen = False
        self._auto_stop_reason = ""
        th = getattr(self, "_battle_watcher", None)
        if th is not None and th.is_alive():
            return
        self._watcher_stop.clear()
        self._battle_watcher = threading.Thread(
            target=self._battle_watcher_loop, daemon=True,
            name="battle-watcher")
        self._battle_watcher.start()

    def _stop_battle_watcher(self) -> None:
        ev = getattr(self, "_watcher_stop", None)
        if ev is not None:
            ev.set()

    def _end_run_session(self, status: str) -> None:
        """结束一次 run 会话：停 F6 服务、写 footer、关日志。

        run 的语义是「会话」：从 run 开始，到 stop（或 quit）结束。
        轴跑完只是到了判断“结束后暂停”的时间点，之后的 F6 暂停、前置
        Excel、重载继续，全都属于同一次 run，日志也要连续写下去。
        """
        self._stop_battle_watcher()
        # 自动 stop（退出关卡）时把真实原因写进 footer，而不是笼统的 stopped
        reason = getattr(self, "_auto_stop_reason", "")
        if status == "stopped" and reason:
            status = reason
        self._auto_stop_reason = ""
        # 轴可能正卡在 _manual_hold 里等第二次 F6；它只在安全点检查
        # stop_requested，所以先登记停止，再停服务循环，否则 join 会超时。
        try:
            if self.executor is not None:
                self.executor.request_stop()
        except Exception:
            pass
        self._stop_f6_service()
        self.running = False
        phase = getattr(self.executor, "phase", "") if self.executor else ""
        footer = (
            f"status: {status}\n"
            f"phase: {phase}\n"
            f"axis: {self.axis_path or '（无）'}\n"
            f"artifacts: out/waa_run.csv, out/pause_log.csv, "
            f"out/run_status.json"
        )
        self.stop_run_log(footer)

    def stop(self) -> None:
        if self.executor is None:
            print("先 load")
            return
        if not self.running:
            print("当前没有运行中的 run")
            return
        # 轴已经跑完（或正在 F6 待命）→ stop 就是结束本次 run：
        # 停 F6 服务、写 footer、关日志。
        if self.f6_service_active or not self.executor.axis_running:
            print("stop：结束本次 run（轴已执行完，日志保存）")
            self._end_run_session("stopped")
            return
        # 轴还在跑：登记停止请求，等当前动作安全退出后由 run 线程收尾
        # （不能清 manual_pause：那会让 _manual_hold 误以为用户按了 F6）。
        self.executor.request_stop()
        print("stop：不会切回游戏窗口；等当前动作安全退出后结束本次 run")

    def pause(self) -> None:
        if self.executor is None:
            print("先 load")
            return
        if not self.executor.manual_pause.is_set():
            self.executor.toggle_manual_pause()
        print("已请求暂停")

    def resume(self) -> None:
        if self.executor is None:
            print("先 load")
            return
        self.executor.manual_pause.clear()
        if not self.running:
            self.executor.resume()
        print("已请求恢复")

    def _preview_status(self) -> dict:
        """没有 load 轴时，status 能给出的预览字段。"""
        c = self.counter
        if c is None:
            return {}
        last = c.last
        in_battle = bool(getattr(c, "in_battle", False))
        rate = getattr(c.counter, "rate_percent", None)
        neg = bool(getattr(c.counter, "neg", False)) and in_battle
        if not in_battle:
            base = getattr(c, "base_rate_percent", PREVIEW_RATE)
            rate_text = f"{float(base):g}%"
        elif rate is None:
            rate_text = ""
        elif neg:
            rate_text = f"{float(rate) / 2:g}%（欠费）"
        else:
            rate_text = f"{float(rate):g}%"
        speed_text = ""
        bullet = False
        if last is not None:
            bullet = bool(getattr(last, "bullet", False))
            try:
                m = float(getattr(last, "m", 1.0) or 1.0)
            except (TypeError, ValueError):
                m = 1.0
            speed_text = "0.2X" if bullet else ("2X" if m >= 1.5 else "1X")
        display_frame = int(c.frame) + int(self.preview_offset)

        # 没读轴时也能读 HUD；OCR 是异步的，等一下最近一次结果。
        hud = {}
        # 只在关卡内做 HUD OCR；主菜单/加载/结算不检测。
        if getattr(c, "in_battle", False):
            try:
                if self.preview_ocr is None:
                    from waa.core.ocr_name import HudOCR
                    self.preview_ocr = HudOCR()
                img = self.stream.latest()
                if img is not None:
                    self.preview_ocr.submit(img)
                    deadline = time.time() + 0.6
                    while time.time() < deadline:
                        hud = self.preview_ocr.values()
                        if hud:
                            break
                        time.sleep(0.05)
            except Exception:
                hud = {}

        cost_text = "" if hud.get("cost") is None else str(hud["cost"])
        deploy_text = ("" if hud.get("deploy_left") is None
                       else str(hud["deploy_left"]))
        selected_text = ""
        stat = {k: "" for k in ("attack", "defense", "resist", "block")}
        if bullet:
            selected_text = hud.get("name") or "?"
            stat = {k: ("" if hud.get(k) is None else str(hud[k]))
                    for k in stat}
        return {
            "status": "预览中（未读轴）",
            "current": "",
            "error": getattr(c.counter, "spec_error", "") or "",
            "frame": str(display_frame),
            "time": _frame_time_text(display_frame),
            "cost_frame": _cost_frame_text(last),
            "rate": rate_text,
            "speed": speed_text,
            "cost": cost_text,
            "deploy_left": deploy_text,
            "selected": selected_text,
            "attack": stat["attack"],
            "defense": stat["defense"],
            "resist": stat["resist"],
            "block": stat["block"],
        }

    @staticmethod
    def _print_run_info(data: dict) -> None:
        print("运行信息：")
        for label, key in (
            ("当前状态", "status"),
            ("当前操作", "current"),
            ("报错信息", "error"),
            ("当前帧号", "frame"),
            ("当前时间", "time"),
            ("费用条", "cost_frame"),
            ("回费速度", "rate"),
            ("当前倍速", "speed"),
            ("当前费用", "cost"),
            ("剩余可放置角色", "deploy_left"),
            ("当前选中", "selected"),
            ("攻击", "attack"),
            ("防御", "defense"),
            ("法抗", "resist"),
            ("阻挡", "block"),
        ):
            value = data.get(key, "")
            if value is None or str(value).strip() == "":
                value = "—"
            print(f"  {label}：{value}")

    def status(self) -> None:
        if self.counter is None:
            print("没有加载轴，也没有预览计数器")
            return
        data = None
        if self.reporter is not None:
            try:
                data = self.reporter.status_snapshot()
            except Exception as exc:
                print(f"读取运行信息失败：{exc}")
        if data is None:
            data = self._preview_status()
        if not data.get("current") and self.executor is not None:
            action = self.executor.current_action
            if action is not None:
                try:
                    data["current"] = str(action[1])
                except Exception:
                    data["current"] = str(action)
        self._print_run_info(data)
        if self.executor is None:
            print("  （未 load 轴：当前为预览模式；HUD 字段按当前画面识别）")

    # ---------- 文件 ----------
    def _pick_file(self, save: bool = False, title: str = "选择 Excel"):
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            os.makedirs(self.default_dir, exist_ok=True)
            if save:
                path = filedialog.asksaveasfilename(
                    title=title, initialdir=self.default_dir,
                    defaultextension=".xlsx",
                    filetypes=[("Excel", "*.xlsx")])
            else:
                path = filedialog.askopenfilename(
                    title=title, initialdir=self.default_dir,
                    filetypes=[("Excel", "*.xlsx")])
            root.destroy()
            return path or None
        except Exception as exc:
            print(f"文件对话框失败：{exc}")
            return None

    def _resolve_path(self, text: str) -> str | None:
        text = text.strip().strip('"').strip("'")
        if not text:
            return None
        if not os.path.isabs(text):
            text = os.path.join(self.default_dir, text)
        return os.path.abspath(text)

    def load_command(self, rest: str) -> bool:
        if self.executor is not None:
            self.executor.mark_suspended = True
        self.reading = True
        ok = False
        try:
            path = self._resolve_path(rest) if rest else self._pick_file(
                title="选择轴 Excel")
            if path:
                ok = self.load_axis(path)
        finally:
            self.reading = False
            if self.executor is not None:
                self.executor.mark_suspended = False
        return ok

    def start_run_log(self, command: str) -> None:
        """每次 run 创建独立日志；stdout/stderr 全线程 tee 到文件。"""
        if self.run_logger is not None:
            return
        try:
            axis_name = os.path.splitext(os.path.basename(
                self.axis_path or "未加载"))[0]
            safe = "".join(
                ch if (ch.isalnum() or ch in "-_") else "_"
                for ch in axis_name)
            stamp = time.strftime("%Y%m%d_%H%M%S")
            log_dir = os.path.join(OUT, "logs")
            os.makedirs(log_dir, exist_ok=True)
            base = os.path.join(log_dir, f"run_{stamp}_{safe}")
            path = base + ".log"
            n = 1
            while os.path.exists(path):
                path = f"{base}_{n}.log"
                n += 1
            header = (
                f"command: {command}\n"
                f"axis: {self.axis_path or '（无）'}\n"
                f"default_dir: {self.default_dir}\n"
                f"screen: {getattr(self.stream, 'cw', '?')}x"
                f"{getattr(self.stream, 'ch', '?')}"
            )
            logger = RunLogger(path, header)
            logger.start()
            self.run_logger = logger
            print(f"运行日志：{path}")
        except Exception as exc:
            self.run_logger = None
            print(f"WARNING: 运行日志创建失败（继续运行）：{exc!r}")

    def stop_run_log(self, footer: str = "") -> None:
        logger = self.run_logger
        self.run_logger = None
        if logger is None:
            return
        logger.stop(footer)
        print(f"运行日志已保存：{logger.path}")

    def log_command(self, rest: str) -> None:
        """log on / off：开关 run 日志记录。默认关闭。"""
        arg = rest.strip().lower()
        if arg in ("on", "1", "开"):
            if self.log_enabled:
                print("日志记录已开启")
            else:
                self.log_enabled = True
                print("日志记录已开启；之后的 run 会写 out\\logs\\run_*.log")
        elif arg in ("off", "0", "关"):
            if not self.log_enabled:
                print("日志记录已关闭")
            else:
                self.log_enabled = False
                print("日志记录已关闭")
        elif arg == "":
            state = "开启" if self.log_enabled else "关闭"
            print(f"日志记录：{state}（用 log on / log off 切换）")
        else:
            print("用法：log on 开启 / log off 关闭 / log 查看状态")

    def verbose_command(self, rest: str) -> None:
        """verbose / debug on/off：开关 run 时的调试细节输出。默认关闭。

        开启后会打印逐帧坐标、选卡重试、子弹进出、步进细节等排查用输出；
        关闭时屏幕只保留动作结果、暂停落点、WARNING/ERROR。
        """
        arg = rest.strip().lower()
        if arg in ("on", "1", "开"):
            if self.executor is not None:
                self.executor.verbose = True
            self._verbose = True
            print("调试输出已开启；run 会打印逐帧/坐标/重试细节")
        elif arg in ("off", "0", "关"):
            if self.executor is not None:
                self.executor.verbose = False
            self._verbose = False
            print("调试输出已关闭")
        elif arg == "":
            state = "开启" if self._verbose else "关闭"
            print(f"调试输出：{state}（用 verbose on / verbose off 切换）")
        else:
            print("用法：verbose on 开启 / verbose off 关闭 / verbose 查看状态")

    def run_command(self, rest: str) -> None:
        """读轴 -> 自动进图（如果在编队界面）-> 开跑。

        不带文件参数时：已经 load 过就直接用当前轴；否则弹文件框。
        """
        if self.running:
            print("上一次 run 还没结束：先输入 stop 结束它（日志会保存），再 run")
            return
        if self.log_enabled:
            self.start_run_log(("run " + rest).strip())
        text = rest.strip()
        if text:
            if not self.load_command(text):
                self.stop_run_log("run 未开始：轴加载失败")
                return
        elif self.executor is None or self.axis is None:
            if not self.load_command(""):
                self.stop_run_log("run 未开始：没有选择轴")
                return
        else:
            print(f"使用已加载的轴：{self.axis_path or '当前轴'}")
        try:
            self._start_execution(auto_enter=True)
        except Exception:
            self.stop_run_log("run 异常退出")
            raise
        if not self.running:
            self.stop_run_log("run 未开始（原因见上方日志）")

    def write_command(self, rest: str) -> None:
        arg = rest.strip()
        if arg.lower() in ("clear", "off", "none"):
            if self.write_writer is not None:
                self.write_writer.close()
            self.write_writer = None
            self.write_path = None
            if self.executor is not None:
                self.executor.mark_writer = None
                self.executor.mark_writer_error = ""
            print("已清除 write：F5 回到 load 轴")
            return
        if arg.lower() in ("new", "create"):
            path = self._pick_file(save=True, title="新建 F5 写入 Excel")
            if not path:
                return
            try:
                shutil.copy2(STANDARD_TEMPLATE, path)
            except Exception as exc:
                print(f"复制标准模板失败：{exc}")
                return
        else:
            path = self._resolve_path(arg) if arg else self._pick_file(
                title="选择 F5 写入 Excel")
            if not path:
                return
        if self.write_writer is not None:
            self.write_writer.close()
            self.write_writer = None
        self.write_path = os.path.abspath(path)
        if not os.path.exists(self.write_path):
            print(f"write 文件不存在：{self.write_path}")
            if self.executor is not None:
                self.executor.mark_writer = None
                self.executor.mark_writer_error = (
                    f"write 文件不存在：{self.write_path}")
            return
        if (self.axis_path is not None
                and _same_path(self.write_path, self.axis_path)):
            print("write = load 轴：F5 写主轴的 Excel")
            if self.executor is not None:
                self.executor.mark_writer = None
                self.executor.mark_writer_error = ""
            return
        try:
            self.write_writer = MarkWriter(self.write_path)
        except Exception as exc:
            print(f"打开 write 文件失败：{exc}")
            self.write_writer = None
            if self.executor is not None:
                self.executor.mark_writer = None
                self.executor.mark_writer_error = (
                    f"write 文件打不开：{exc}")
            return
        if self.executor is not None:
            self.executor.mark_writer = self.write_writer
            self.executor.mark_writer_error = ""
        print(f"write = {self.write_path}")

    def bind_command(self, rest: str) -> None:
        if not rest:
            print(f"默认文件夹：{self.default_dir}")
            return
        path = os.path.abspath(os.path.expanduser(rest.strip().strip('"')))
        if not os.path.isdir(path):
            print(f"不是文件夹：{path}")
            return
        self.default_dir = path
        self.config["default_dir"] = path
        _save_config(self.config)
        print(f"默认文件夹：{path}")

    def cards_command(self) -> None:
        """识别当前待部署区的每一张卡是谁。"""
        frame = self.stream.latest()
        if frame is None:
            print("还没有抓到游戏画面，请确认游戏窗口正常")
            return
        print("识别待部署区…")
        try:
            cards = identify_cards(frame, log=print)
        except Exception as exc:
            print(f"待部署区识别失败：{exc}")
            return
        if not cards:
            print("没有找到待部署卡")
            return
        print(f"待部署区：检测到 {len(cards)} 张卡")
        for i, card in enumerate(cards, 1):
            name = card.get("oper")
            score = card.get("identify_score", 0.0)
            if name is None:
                name = f"未识别（候选 {card.get('candidate') or '—'}）"
            x, y = card.get("avatar_center", card["click"])
            print(f"  {i:2d}. {name:<12s} 分数={score:.3f} "
                  f"头像中心=({x:.1f},{y:.1f})")
        right = max(cards, key=lambda c: c.get(
            "avatar_center", c["click"])[0])
        rx, ry = right.get("avatar_center", right["click"])
        print(f"最右卡：{right.get('oper') or right.get('candidate') or '—'} "
              f"头像中心=({rx:.1f},{ry:.1f})")
        print("说明：位置是 1280×720 标准坐标；同一张卡可能显示"
              "剩余使用次数（如 ×2），仍算一张卡。")

    def showframe_command(self, rest: str) -> None:
        text = rest.strip()
        if text.lower().startswith("auto"):
            arg = text[4:].strip().lower()
            if arg in ("on", "off"):
                self.auto_showframe = (arg == "on")
                self.config["showframe_auto"] = self.auto_showframe
                _save_config(self.config)
            elif arg:
                print("用法：showframe auto on / showframe auto off")
                return
            state = "开" if self.auto_showframe else "关"
            print(f"启动时自动打开帧窗：{state}"
                  "（showframe auto on/off 切换）")
            return
        if self.counter is None:
            self._start_preview_counter()
        if text:
            alpha = _parse_alpha(text)
            if alpha is None:
                return
            self.showframe.set_alpha(alpha)
            self.config["showframe_alpha"] = self.showframe.alpha
            _save_config(self.config)
        self.showframe.show()
        print(f"帧窗口已打开，透明度 {round(self.showframe.alpha * 100)}%")
        if self.executor is None:
            print("当前未 load 轴：显示的是预览帧号"
                  f"（{PREVIEW_RATE:g}%，自动识别 50/110）")

    # ---------- 主循环 ----------
    def run_console(self) -> int:
        # 头像库预热是后台线程：缓存命中时约 0.2 秒。这里先等它一下，让
        # "[vision] 头像库缓存已加载" 出现在提示符之前，而不是被 _PromptGuard
        # 扣到第一次回车之后（那样看起来像"敲了命令才开始加载"）。
        preload = getattr(self, "_card_preload_thread", None)
        if preload is not None and preload.is_alive():
            preload.join(timeout=1.5)
        print(f"WAA (WrankwoodAssistantArknights) v{__version__}（{VERSION_CHANNEL}）")
        print("键位：F5 打轴记录当前帧 · F6 暂停/恢复")
        print("输入 help 查看全部命令；log on 开启日志记录。")
        if self.auto_showframe and not getattr(self.args, "no_showframe", False):
            try:
                if self.counter is None:
                    self._start_preview_counter()
                self.showframe.show()
            except Exception as exc:
                print(f"帧窗口自动打开失败：{exc!r}"
                      "（用 showframe 手动开，或 showframe auto off 关掉自动打开）")
        guard = _PromptGuard(sys.stdout)
        sys.stdout = guard
        try:
            return self._console_loop(guard)
        finally:
            guard.release()
            if sys.stdout is guard:
                sys.stdout = guard._stream

    def _console_loop(self, guard: _PromptGuard) -> int:
        while True:
            guard.hold()
            try:
                line = input("waa> ").strip()
            except (EOFError, KeyboardInterrupt):
                print()          # 先把光标换到下一行，再吐积压的后台输出
                guard.release()
                self.quit()
                return 0
            finally:
                guard.release()
            if not line:
                continue
            if self.run_logger is not None:
                self.run_logger.command(line)
            parts = line.split(maxsplit=1)
            cmd = parts[0].lower()
            rest = parts[1] if len(parts) > 1 else ""
            try:
                if cmd in ("quit", "exit", "q"):
                    self.quit()
                    return 0
                if cmd == "help":
                    self.print_help()
                elif cmd == "load":
                    self.load_command(rest)
                elif cmd == "read":
                    print("read 已改名为 load")
                    self.load_command(rest)
                elif cmd == "run":
                    self.run_command(rest)
                elif cmd == "log":
                    self.log_command(rest)
                elif cmd in ("verbose", "debug"):
                    self.verbose_command(rest)
                elif cmd == "write":
                    self.write_command(rest)
                elif cmd == "bind":
                    self.bind_command(rest)
                elif cmd == "update":
                    self.update_command()
                elif cmd == "showframe":
                    self.showframe_command(rest)
                elif cmd in ("cards", "units"):
                    self.cards_command()
                elif cmd in ("hideframe", "closeframe"):
                    self.showframe.hide()
                    print("帧窗口已关闭")
                elif cmd == "start":
                    print("start 已移除；请用 run")
                elif cmd == "stop":
                    self.stop()
                elif cmd == "pause":
                    self.pause()
                elif cmd == "resume":
                    self.resume()
                elif cmd == "status":
                    self.status()
                elif cmd == "offset":
                    self.offset_command(rest)
                elif cmd == "frame":
                    self.frame_command(rest)
                else:
                    print(f"未知命令：{cmd}")
                    self.print_help()
            except Exception as exc:
                print(f"[错误] {cmd} 失败：{exc}")

    def offset_command(self, rest: str) -> None:
        s = rest.strip()
        if not s:
            if self.executor is not None:
                print(f"offset = {self.executor.start_offset}")
            else:
                print(f"offset = {self.preview_offset}")
            return
        try:
            value = int(s)
        except ValueError:
            print("offset 请输入整数，例如 offset 5 / offset +5 / offset -5")
            return
        if self.executor is not None:
            # offset 永远是绝对值；+5/-5 只是允许带符号。
            self.executor.set_offset(value=value)
            self.preview_offset = int(self.executor.start_offset)
        else:
            self.preview_offset = value
            cur = (self.counter.frame + self.preview_offset
                   if self.counter is not None else None)
            print(f"offset = {value}"
                  + (f"（当前帧 {cur}）" if cur is not None else ""))

    def frame_command(self, rest: str) -> None:
        s = rest.strip()
        if not s:
            if self.executor is not None:
                print(f"当前帧 = {self.executor.current_frame()}")
            elif self.counter is not None:
                print(f"当前帧 = {self.counter.frame + self.preview_offset}")
            else:
                print("没有帧计数器")
            return
        try:
            value = int(s)
        except ValueError:
            print("frame 请输入整数，例如 frame 500 / frame +5 / frame -5")
            return
        relative = s[0] in "+-"
        if self.executor is not None:
            if relative:
                self.executor.set_offset(delta=value)
            else:
                self.executor.set_frame(value)
            self.preview_offset = int(self.executor.start_offset)
            return
        if self.counter is None:
            print("没有帧计数器")
            return
        if relative:
            self.preview_offset += value
        else:
            self.preview_offset = value - int(self.counter.frame)
        print(f"frame = {self.counter.frame + self.preview_offset}"
              f"（offset = {self.preview_offset}）")

    @staticmethod
    def print_help() -> None:
        print("命令：")
        print("")
        print("  ── 操作 ──")
        print("  load [文件]      读取轴并打开 Excel 面板（无参数弹文件框）")
        print("  run [文件]       运行：读轴/复用已加载轴 → 自动进图 → 开始")
        print("  stop             停止执行")
        print("  pause / resume   手动暂停 / 恢复")
        print("  write [文件]     选择 F5 写入 Excel；write new 新建；write clear 清除")
        print("")
        print("  ── 查看 ──")
        print("  status           输出运行信息全部字段（未 load 也能看预览）")
        print("  cards / units    识别待部署区每张卡是谁及头像中心")
        print("  frame [值|+n|-n] 查看/设置当前帧；+n/-n 相对")
        print("  offset [值]      查看/设置 offset（绝对；可写 +5/-5）")
        print("")
        print("  ── 系统 ──")
        print("  log on / off     开启 / 关闭 run 日志记录（默认关闭）")
        print("  verbose on / off 开启 / 关闭调试细节输出（默认关闭）")
        print("  bind [文件夹]    设置/查看默认文件夹")
        print("  update           立即从 GitHub 更新本地地图/头像/角色表")
        print("  showframe [10-100] 打开三行帧窗口并设透明度；不用先 load")
        print("  showframe auto on/off 启动时是否自动打开帧窗（默认开）")
        print("  hideframe        关闭三行帧窗口")
        print("  quit             退出")

    def quit(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._stop_battle_watcher()
        if self.running:
            self._end_run_session("quit")
        else:
            self.stop_run_log("控制台退出")
        try:
            self.showframe.stop()
        except Exception:
            pass
        self._teardown_runtime()
        if self.write_writer is not None:
            self.write_writer.close()
        try:
            self.stream.stop()
        except Exception:
            pass
        try:
            self.touch.close()
        except Exception:
            pass
        print("bye")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="WAA 交互式控制台")
    p.add_argument("--axis", default=None,
                   help="启动时直接读取的轴 Excel；不填则进入控制台等 load/run")
    p.add_argument("--start-action", type=int, default=0)
    p.add_argument("--start-frame", type=int, default=None)
    p.add_argument("--start-offset", type=int, default=0)
    p.add_argument("--pause-mode",
                   choices=["tap", "hold", "esc", "center", "key_first",
                            "settle", "direct"], default="hold")
    p.add_argument("--speed", choices=["auto", "1", "2"], default="auto")
    p.add_argument("--no-step", action="store_true")
    p.add_argument("--pause-method", choices=["touch", "esc"],
                   default="touch")
    p.add_argument("--cost-rate", type=float, default=None)
    p.add_argument("--negative", choices=["auto", "on", "off"],
                   default="auto")
    p.add_argument("--no-excel", action="store_true")
    p.add_argument("--no-update", action="store_true",
                   help="启动时不检查 GitHub 数据更新")
    p.add_argument("--copy", action="store_true",
                   help="复制轴文件再写面板；默认原地写回")
    p.add_argument("--selftest", action="store_true",
                   help="自检：验证依赖/资源/写权限，不碰游戏也不抢单实例锁")
    p.add_argument("--no-showframe", action="store_true",
                   help="启动时不自动打开悬浮帧窗（只影响这一次）")
    return p


# 单实例锁：命名 Mutex 由系统随进程退出自动释放，不会像锁文件那样残留。
_SINGLE_INSTANCE_NAME = "Local\\WAA_AXIS_CONSOLE_V1"
_SINGLE_INSTANCE_HANDLE = None
# 控制台窗口标题：第二个实例靠它把已经在跑的那个窗口拉到前台。
CONSOLE_TITLE = "WAA 控制台"


def _set_console_title() -> None:
    try:
        ctypes.windll.kernel32.SetConsoleTitleW(CONSOLE_TITLE)
    except Exception:
        pass


def _console_windows() -> list[int]:
    """所有标题像 WAA 控制台的可见顶层窗口（排除自己这个）。"""
    user32 = ctypes.windll.user32
    k32 = ctypes.windll.kernel32
    user32.FindWindowW.restype = ctypes.c_void_p
    user32.GetWindowTextLengthW.argtypes = (ctypes.c_void_p,)
    user32.GetWindowTextW.argtypes = (ctypes.c_void_p, ctypes.c_wchar_p,
                                      ctypes.c_int)
    user32.IsWindowVisible.argtypes = (ctypes.c_void_p,)
    k32.GetConsoleWindow.restype = ctypes.c_void_p

    # 排除自己：ConPTY 下 GetConsoleWindow 是伪窗口，真实窗口在 owner 上，
    # 用 _console_hwnd() 才能拿到自己真正的那一个，否则会把自己当成
    # “已经运行的另一个控制台”并把自己拉到前台。
    own = _console_hwnd() or k32.GetConsoleWindow()
    found: list[int] = []

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p,
                                     ctypes.c_void_p)
    user32.EnumWindows.argtypes = (WNDENUMPROC, ctypes.c_void_p)

    def _visit(hwnd, _lparam):
        if hwnd == own or not user32.IsWindowVisible(hwnd):
            return True
        n = int(user32.GetWindowTextLengthW(hwnd))
        if n <= 0:
            return True
        buf = ctypes.create_unicode_buffer(n + 1)
        user32.GetWindowTextW(hwnd, buf, n + 1)
        text = buf.value.strip()
        # 只认自己设的标题：弱匹配（含 waa_console/axis_console）会把
        # “自己这个新窗口”或别的 cmd 也认进来，反而不安全。
        if text == CONSOLE_TITLE or text.startswith(CONSOLE_TITLE):
            found.append(hwnd)
        return True

    user32.EnumWindows(WNDENUMPROC(_visit), 0)
    return found


def _is_foreground(hwnd: int) -> bool:
    """前台窗口是不是这个窗口（把子窗口归一到顶层再比）。

    Unity 客户端有时会把焦点放在子窗口上，直接比 hwnd 会误判。
    """
    user32 = ctypes.windll.user32
    user32.GetForegroundWindow.restype = ctypes.c_void_p
    user32.GetAncestor.argtypes = (ctypes.c_void_p, ctypes.c_uint)
    user32.GetAncestor.restype = ctypes.c_void_p
    fg = int(user32.GetForegroundWindow() or 0)
    if fg == 0:
        return False
    return int(user32.GetAncestor(fg, 2) or fg) == int(hwnd) or fg == int(hwnd)


def _focus_window(hwnd: int, timeout: float = 0.6) -> bool:
    """还原并前置指定窗口，**确认真的到前台**才返回 True。

    SetForegroundWindow 的返回值不可信：它只表示调用被接受，而前台切换是
    异步的（实测 30~50ms 才生效），被前台锁挡住时也可能返回非 0。
    2026-11-09 06:09 那轮就是「返回 True → 立刻 GetForegroundWindow → 还是
    旧窗口 → 判定不在前台 → 取消自动进图」，一次都没点就退出了。
    所以这里调用后轮询验证，并用 AttachThreadInput / ALT 键兜底。
    """
    user32 = ctypes.windll.user32
    k32 = ctypes.windll.kernel32
    user32.ShowWindow.argtypes = (ctypes.c_void_p, ctypes.c_int)
    user32.SetForegroundWindow.argtypes = (ctypes.c_void_p,)
    user32.BringWindowToTop.argtypes = (ctypes.c_void_p,)
    user32.GetForegroundWindow.restype = ctypes.c_void_p
    user32.GetWindowThreadProcessId.argtypes = (ctypes.c_void_p,
                                                ctypes.POINTER(ctypes.c_ulong))
    user32.AttachThreadInput.argtypes = (ctypes.c_ulong, ctypes.c_ulong,
                                         ctypes.c_int)

    def settle(deadline: float) -> bool:
        while time.time() < deadline:
            if _is_foreground(hwnd):
                return True
            time.sleep(0.01)
        return _is_foreground(hwnd)

    user32.ShowWindow(hwnd, 9)          # SW_RESTORE
    try:
        user32.BringWindowToTop(hwnd)
    except Exception:
        pass
    deadline = time.time() + timeout

    # 1) 直接抢
    user32.SetForegroundWindow(hwnd)
    if settle(deadline):
        return True

    # 2) 附到当前前台线程再抢（绕过前台锁）
    try:
        fg = int(user32.GetForegroundWindow() or 0)
        tid_fg = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(fg, ctypes.byref(tid_fg))
        tid_me = k32.GetCurrentThreadId()
        user32.AttachThreadInput(tid_me, tid_fg.value, 1)
        try:
            user32.SetForegroundWindow(hwnd)
        finally:
            user32.AttachThreadInput(tid_me, tid_fg.value, 0)
        if settle(min(deadline, time.time() + 0.25)):
            return True
    except Exception:
        pass

    # 3) 模拟一次 ALT 按键解除前台锁，再抢一次
    try:
        VK_MENU = 0x12
        KEYEVENTF_KEYUP = 0x0002
        user32.keybd_event(VK_MENU, 0, 0, 0)
        user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.02)
        user32.SetForegroundWindow(hwnd)
    except Exception:
        pass
    return _is_foreground(hwnd)


def _console_hwnd() -> int:
    """本进程**真正可见**的控制台窗口句柄；拿不到返回 0。

    Windows 11 的 ConPTY 下 GetConsoleWindow() 返回的是
    `PseudoConsoleWindow`——一个无标题、不可聚焦的伪窗口，真正的窗口挂在
    它的 owner 链上。直接拿伪窗口去 SetForegroundWindow，会把真正的窗口
    带出来但 `_is_foreground` 永远不成立（2026-11-09 06:1x 实测）。
    所以沿 owner 链向上找第一个「可见且标题非空」的窗口。
    """
    try:
        h = int(ctypes.windll.kernel32.GetConsoleWindow())
    except Exception:
        return 0
    if not h:
        return 0
    seen: set[int] = set()
    cur = h
    while cur and cur not in seen:
        seen.add(cur)
        try:
            if win32gui.IsWindowVisible(cur) and win32gui.GetWindowText(cur):
                return cur
        except Exception:
            return 0
        try:
            nxt = win32gui.GetWindow(cur, 4)          # GW_OWNER
        except Exception:
            nxt = 0
        if not nxt:
            try:
                root = win32gui.GetAncestor(cur, 2)   # GA_ROOT
            except Exception:
                root = 0
            nxt = 0 if root == cur else root
        cur = nxt
    try:
        return h if win32gui.IsWindowVisible(h) else 0
    except Exception:
        return 0


def _bring_console_to_front() -> bool:
    """把本进程的控制台窗口拉到前台（报错停止时用，让 traceback 看得见）。"""
    hwnd = _console_hwnd()
    if not hwnd:
        return False
    return _focus_window(hwnd)


def _focus_existing_console() -> bool:
    """把已经在跑的那个 WAA 控制台窗口拉到前台。"""
    for hwnd in _console_windows():
        if _focus_window(hwnd):
            return True
    return False


def _acquire_single_instance() -> bool:
    """拿到单实例锁；已有控制台在跑时返回 False。

    必须在 ConsoleSession 之前检查：第二个实例不会去抓屏、注册 F5/F6 热键、
    打开 Excel，也不会覆盖 out/ 下的日志、CSV 和 run_status.json。
    """
    global _SINGLE_INSTANCE_HANDLE
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateMutexW.argtypes = (ctypes.c_void_p, ctypes.c_int,
                                 ctypes.c_wchar_p)
    k32.CreateMutexW.restype = ctypes.c_void_p
    k32.CloseHandle.argtypes = (ctypes.c_void_p,)
    k32.CloseHandle.restype = ctypes.c_int
    ERROR_ALREADY_EXISTS = 183
    for _ in range(3):
        handle = k32.CreateMutexW(None, 0, _SINGLE_INSTANCE_NAME)
        err = ctypes.get_last_error()
        if handle and err != ERROR_ALREADY_EXISTS:
            _SINGLE_INSTANCE_HANDLE = handle   # 持有句柄，别让 GC 提前关掉
            return True
        if handle:
            k32.CloseHandle(handle)
        # 上一个实例可能正在退出，稍等再确认，避免误判。
        time.sleep(0.3)
    return False


def run_selftest() -> int:
    """自检：只验证「这个包在当前机器上能不能跑」，不碰游戏、不抢单实例锁。"""
    import importlib

    ok = True

    def check(name, fn):
        nonlocal ok
        try:
            detail = fn()
            print(f"PASS {name}" + (f"  {detail}" if detail else ""))
        except Exception as exc:
            ok = False
            print(f"FAIL {name}  {exc!r}")

    def _env():
        frozen = getattr(sys, "frozen", False)
        return (f"frozen={frozen}  app_root={app_root()}  "
                f"bundle={bundle_root()}")

    def _imports():
        mods = ["cv2", "numpy", "openpyxl", "PIL.Image", "win32gui",
                "win32api", "pythoncom", "win32com.client", "windows_capture",
                "winrt.windows.media.ocr", "winrt.windows.globalization",
                "winrt.windows.graphics.imaging", "winrt.windows.storage",
                "winrt.windows.storage.streams"]
        bad = []
        for name in mods:
            try:
                importlib.import_module(name)
            except Exception as exc:
                bad.append(f"{name}({exc})")
        if bad:
            raise RuntimeError("；".join(bad))
        return f"{len(mods)} 个模块"

    def _assets():
        paths = [asset_path("waa", "resources", "BattleOfficiallyBegin.png"),
                 asset_path("waa", "resources", "BattleOpersFlag.png"),
                 asset_path("waa", "resources", "StartOperation.png"),
                 asset_path("waa", "templates", "WAA_标准模板.xlsx")]
        missing = [p for p in paths if not os.path.exists(p)]
        if missing:
            raise FileNotFoundError(missing[0])
        return f"{len(paths)} 个文件"

    def _out_writable():
        os.makedirs(OUT, exist_ok=True)
        probe = os.path.join(OUT, "_selftest.tmp")
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("ok")
        os.remove(probe)
        return OUT

    def _frame_exe():
        path = os.path.join(ROOT, "WAA_frame.exe")
        if getattr(sys, "frozen", False) and not os.path.exists(path):
            raise FileNotFoundError(f"缺少帧窗程序 {path}")
        return "帧窗程序就位" if os.path.exists(path) else "开发模式，用 pythonw"

    def _tk_dialog():
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        root.destroy()
        return "文件对话框可用"

    def _ocr():
        from winrt.windows.globalization import Language
        from winrt.windows.media.ocr import OcrEngine
        engine = (OcrEngine.try_create_from_language(Language("zh-Hans"))
                  or OcrEngine.try_create_from_user_profile_languages())
        if engine is None:
            raise RuntimeError("没有可用的 OCR 语言包（需中文 OCR）")
        return "OCR 引擎可用"

    print(f"WAA 自检  v{__version__}（{VERSION_CHANNEL}）")
    check("运行环境", _env)
    check("第三方依赖", _imports)
    check("随包资源", _assets)
    check("out 可写", _out_writable)
    check("帧窗程序", _frame_exe)
    check("文件对话框", _tk_dialog)
    check("OCR 语言包", _ocr)
    print("结果：" + ("全部通过" if ok else "有失败项，见上面 FAIL 行"))
    return 0 if ok else 1


def _force_utf8_stdio() -> None:
    """输出被重定向时也按 UTF-8 写，避免日志/管道里出现乱码。

    连着真实控制台时 stdout 本来就是 UTF-8（走 WriteConsoleW），这里等于空操作。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return True


def _relaunch_as_admin() -> bool:
    """以管理员身份重启自己；返回 True 表示已发起提权，当前进程该退出了。

    打包后不再用 .bat 提权（少一个文件、路径不受工作目录影响）：
    双击 WAA.exe -> UAC -> 新的管理员进程接管，原进程立刻退出。
    """
    if getattr(sys, "frozen", False):
        exe = sys.executable
        params = subprocess.list2cmdline(sys.argv[1:])
    else:
        exe = sys.executable
        params = subprocess.list2cmdline(
            [os.path.abspath(__file__)] + sys.argv[1:])
    try:
        rc = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", exe, params, os.getcwd(), 1)
    except Exception:
        return False
    return rc > 32


def _ensure_admin(args) -> int | None:
    """需要管理员时自动提权。返回 None 表示继续跑，返回整数表示直接退出。"""
    if _is_admin():
        return None
    if os.environ.get("WAA_NO_ELEVATE") == "1":
        return None          # 自动化测试用：跳过提权
    if getattr(sys, "frozen", False):
        print("需要管理员权限：正在请求提权（在 UAC 窗口点“是”）…")
        if _relaunch_as_admin():
            return 0
        print("提权被取消。触摸注入不会生效，请右键 WAA.exe → 以管理员身份运行。")
        if sys.stdin is not None and sys.stdin.isatty():
            input("按回车退出…")
        return 1
    print("[提示] 当前不是管理员：触摸注入会失效。建议在管理员终端里运行。")
    return None


def main() -> int:
    _force_utf8_stdio()
    args = build_parser().parse_args()
    if getattr(args, "selftest", False):
        return run_selftest()
    code = _ensure_admin(args)
    if code is not None:
        return code
    if not _acquire_single_instance():
        # 已经有实例在跑：把它拉到前台，自己安静退出（退出码 2 = 让 bat 不要 pause）
        if _focus_existing_console():
            return 2
        print("已有一个 WAA 控制台在运行，但没找到它的窗口。")
        print("请在任务栏里找标题是“WAA 控制台”的那个窗口；"
              "要重开请先在那里输入 quit。")
        return 1
    _set_console_title()
    try:
        session = ConsoleSession(args)
    except Exception as exc:
        print(f"[错误] 启动失败：{exc}")
        _bring_console_to_front()
        return 1
    try:
        if args.axis:
            session.load_axis(args.axis)
        return session.run_console()
    except KeyboardInterrupt:
        # 用户自己按的 Ctrl+C，窗口本来就在前台，不用切。
        print("\n已中断")
        return 130
    except SystemExit:
        raise
    except BaseException as exc:
        print(f"\n[错误] 控制台异常退出：{exc}")
        _bring_console_to_front()
        return 1
    finally:
        try:
            session.quit()
        except Exception as exc:
            # quit 里再抛异常会盖掉上面真正的报错，单独打出来。
            print(f"[错误] 退出清理失败：{exc}")


if __name__ == "__main__":
    sys.exit(main())
