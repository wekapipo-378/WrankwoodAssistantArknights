"""运行轴（第一版）。

管理员权限运行：
    Start-Process python -ArgumentList "tools\\run_axis.py --axis ... --max-actions 1" -Verb RunAs
"""

from __future__ import annotations

import argparse
import ctypes
import os
import sys
import threading
import time
import traceback
from ctypes import wintypes

import cv2
import win32gui

if getattr(sys, "frozen", False):
    ROOT = os.path.dirname(os.path.abspath(sys.executable))
else:
    ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from waa.core.waa_counter import WAACounter, FrameCounterThread  # noqa: E402
from waa.core.axis import read_axis  # noqa: E402
from waa.core.cost_model import (  # noqa: E402
    SPEED, parse_declared_speeds, rate_from_frames, rate_to_frames,
)
from waa.core.capture import grab_client  # noqa: E402
from waa.core.capture_stream import CaptureStream  # noqa: E402
from waa.core.executor import Executor, FatalReloadError  # noqa: E402
from waa.core.excel_report import ExcelReporter  # noqa: E402
from waa.core.send_input import click as mouse_click  # noqa: E402
from waa.core.touch_input import TouchDevice  # noqa: E402
from PIL import Image  # noqa: E402

OUT = os.path.join(ROOT, "out")


def _resolve_rate_config(axis, override):
    """回费速度来源（按优先级）：

    1. 命令行 --cost-rate（调试用）
    2. 轴设置“可出现回费速度”：按顺序写本局可能出现的正常回费速度，
       例如 1 / 1.1 / 0.5（100%、110%、50%）。第一格是基准。
    3. 旧设置“费用回复率/回费率”或“每费帧数”（30=100%，60=50%）
    4. 默认 100%

    负费固定按当前正常速度 ÷ 2，不读第二档。
    返回 (base_rate, neg_speed, neg_declared, source, declared)。
    """
    declared_key = None
    raw_declared = None
    for key in ("回费速度", "可出现回费速度", "可能出现回费速度"):
        if axis.settings.get(key) is not None:
            declared_key = key
            raw_declared = axis.settings.get(key)
            break
    declared = parse_declared_speeds(raw_declared)
    neg_declared = True   # 负费固定为当前速度的一半

    def normal_speed(base: float) -> int:
        return max(1, int(round(SPEED * base / 100.0)))

    def neg_for(base: float) -> int:
        return max(1, normal_speed(base) // 2)

    if override is not None:
        base = float(override)
        return (base, neg_for(base), True, "--cost-rate", declared)

    if declared:
        base = declared[0]
        return (base, neg_for(base), neg_declared,
                declared_key or "回费速度", declared)

    for key in ("费用回复率", "回费率", "回复率"):
        v = axis.settings.get(key)
        if v is None or str(v).strip() == "":
            continue
        try:
            base = float(str(v).strip().rstrip("%"))
            return (base, neg_for(base), True, key, declared)
        except ValueError:
            print(f"设置 {key}={v!r} 不是数字，已忽略")
    v = axis.settings.get("每费帧数")
    if v is not None and str(v).strip() != "":
        try:
            base = rate_from_frames(float(str(v).strip()))
            return (base, neg_for(base), True, "每费帧数", declared)
        except ValueError:
            print(f"设置 每费帧数={v!r} 不是数字，已忽略")
    return 100.0, neg_for(100.0), True, "默认", declared


WM_HOTKEY = 0x0312
MOD_NOREPEAT = 0x4000
VK_F5 = 0x74
VK_F6 = 0x75


def start_pause_hotkey(executor) -> threading.Thread:
    """F5 打轴、F6 手动暂停 / 恢复；被占用时只提示，不改用其他键。"""

    def worker() -> None:
        u = ctypes.windll.user32
        f6_ok = u.RegisterHotKey(None, 6, MOD_NOREPEAT, VK_F6)
        f5_ok = u.RegisterHotKey(None, 5, MOD_NOREPEAT, VK_F5)
        if not f6_ok:
            print("[hotkey] F6 注册失败：可能被其他程序占用；"
                  "手动暂停本次不可用")
        if not f5_ok:
            print("[hotkey] F5 注册失败：可能被其他程序占用；"
                  "打轴取帧本次不可用")
        if f6_ok:
            print("[hotkey] F6 就绪：手动暂停 / 恢复")
        if f5_ok:
            print("[hotkey] F5 就绪：在作战记录末尾记录当前帧")
        if not (f5_ok or f6_ok):
            return
        msg = wintypes.MSG()
        while u.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            if msg.message != WM_HOTKEY:
                continue
            if msg.wParam == 6:
                executor.toggle_manual_pause()
            elif msg.wParam == 5:
                executor.mark_frame()

    t = threading.Thread(target=worker, daemon=True)
    t.start()
    return t


def _setting_bool(value, default: bool = False) -> bool:
    """y/n、是/否、开/关、1/0 之类转布尔。"""
    s = str(value).strip().lower() if value is not None else ""
    if s in ("y", "yes", "是", "开", "on", "true", "1"):
        return True
    if s in ("n", "no", "否", "关", "off", "false", "0"):
        return False
    return default


def _focus_window(hwnd: int) -> bool:
    """把游戏窗口拉到前台；打开 Excel 后必须再调一次。"""
    if win32gui.GetForegroundWindow() == hwnd:
        return True
    try:
        win32gui.ShowWindow(hwnd, 6)   # SW_MINIMIZE
        time.sleep(0.15)
        win32gui.ShowWindow(hwnd, 9)   # SW_RESTORE
        time.sleep(0.25)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    for _ in range(20):
        if win32gui.GetForegroundWindow() == hwnd:
            return True
        time.sleep(0.05)
    return False


def _post_run_manual_pause(executor, reporter) -> None:
    """轴跑完后程序不退出；F6 仍可暂停、前置 Excel，再按 F6 恢复。"""
    executor.axis_running = False
    print("轴已结束：F6 手动暂停")
    if executor.is_running_icon():
        executor.pause(assume_running=True)
    if (getattr(reporter, "enabled", False)
            and not getattr(reporter, "started", False)):
        print("正在打开 Excel 面板…")
        reporter.start(wait_ready=True)
    if getattr(reporter, "started", False):
        reporter.bring_to_front()
    while executor.manual_pause.is_set():
        time.sleep(0.05)
    executor._focus_game()
    executor._park_cursor()
    executor.resume()
    print("轴已结束：F6 恢复")


def _resolve_negative(axis, override: str):
    """局内负费：auto=看减号，on/off=强制。返回 (True/False/None, 来源)。"""
    if override != "auto":
        return override == "on", f"--negative {override}"
    v = axis.settings.get("局内负费")
    if v is not None:
        s = str(v).strip().lower()
        if s in ("开", "强制开", "on", "true", "1", "yes"):
            return True, "轴设置 局内负费"
        if s in ("关", "强制关", "off", "false", "0", "no"):
            return False, "轴设置 局内负费"
    return None, "自动（减号识别）"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--axis", default=os.path.join(
        ROOT, "axes", "WAA_sample_1-7.xlsx"))
    parser.add_argument("--max-actions", type=int, default=1)
    parser.add_argument("--start-action", type=int, default=0)
    parser.add_argument("--start-frame", type=int, default=None,
                        help="强制当前全局帧（用于从关卡中途继续测试）")
    parser.add_argument("--pause-mode",
                        choices=["tap", "hold", "esc", "center", "key_first", "settle", "direct"],
                        default="hold", help="技能/撤退时的选中方式；direct=暂停中直接点干员")
    parser.add_argument("--speed", choices=["auto", "1", "2"], default="auto",
                        help="倍速；auto 从 WAA 读数判断")
    parser.add_argument("--no-step", action="store_true",
                        help="关闭单帧步进（调试用）")
    parser.add_argument("--pause-method", choices=["touch", "esc"], default="touch",
                        help="暂停方式：touch 右上角按钮，esc 键盘 ESC")
    parser.add_argument("--start-stage", action="store_true",
                        help="在编队界面点“开始行动”，等战斗开始后跑轴")
    parser.add_argument("--cost-rate", type=float, default=None,
                        help="基准回费速度百分比（调试用，100=30帧/费）；"
                             "不填则读轴设置“可出现回费速度”")
    parser.add_argument("--negative", choices=["auto", "on", "off"],
                        default="auto",
                        help="局内负费：auto=看费用数字左边的减号，on/off=强制")
    parser.add_argument("--no-excel", action="store_true",
                        help="不打开 Excel 运行面板（调试用）")
    parser.add_argument("--in-place", action="store_true",
                        help="直接写回输入的 xlsx（原地模式）；运行前请先关闭该文件")
    args = parser.parse_args()

    axis = read_axis(args.axis)
    print(f"axis={args.axis} map={axis.map_code} actions={len(axis.actions)} warnings={axis.warnings}")
    for a in axis.actions:
        print(" ", a)

    rate, neg_speed, neg_declared, rate_src, declared = _resolve_rate_config(
        axis, args.cost_rate)
    frames_per_cost = rate_to_frames(rate)
    if declared:
        print("回费速度 = "
              + " / ".join(f"{x:g}%" for x in declared))
    print(f"基准回费速度 = {rate:g}%（一费 {frames_per_cost:.2f} 逻辑帧），"
          f"来源：{rate_src}")
    if rate > 0:
        neg_pct = rate / 2.0
        print(f"  局内负费 = {rate:g}% ÷ 2 = {neg_pct:g}%"
              f"（一费 {rate_to_frames(neg_pct):.2f} 逻辑帧）")
    if rate <= 0:
        print("回费率为 0：费用条完全不增长，纯视觉方案没有任何帧信息，已停止。")
        return 3
    slowest = frames_per_cost
    if rate > 0:
        slowest = max(slowest, rate_to_frames(rate / 2.0))
    if declared:
        slowest = max(slowest, max(rate_to_frames(x) for x in declared))
        slowest = max(slowest, rate_to_frames(min(declared) / 2.0))
    if slowest > 187.0:
        print("提示：有效回费速度低于约 16%，条宽会连续多帧不变；计数器要等"
              "宽度变化时才能补上这几帧，打轴落点可能受影响。")
    neg_override, neg_src = _resolve_negative(axis, args.negative)
    print(f"局内负费 = {neg_src}")
    default_2x = _setting_bool(axis.settings.get("默认2倍速"), False)
    start_pause = _setting_bool(axis.settings.get("开始前暂停"), False)
    end_pause = _setting_bool(axis.settings.get("结束后暂停"), False)
    print(f"轴设置：默认2倍速={default_2x} 开始前暂停={start_pause} "
          f"结束后暂停={end_pause}")

    hwnd = win32gui.FindWindow("UnityWndClass", "明日方舟")
    if not hwnd:
        raise RuntimeError("找不到游戏窗口")
    if ctypes.windll.user32.IsZoomed(hwnd):
        win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
    # 单纯 SetForegroundWindow 会被 Windows 的前台锁挡掉；用最小化/还原
    # 强制把游戏放到前台，否则 ESC 和触控都会送给别的窗口。
    _focus_window(hwnd)
    time.sleep(0.3)
    print(f"foreground={win32gui.GetForegroundWindow() == hwnd} client={win32gui.GetClientRect(hwnd)}")

    Image.fromarray(grab_client(hwnd)[:, :, ::-1]).save(os.path.join(OUT, "axis_before.png"))

    touch = TouchDevice()
    stream = CaptureStream(hwnd)
    stream.start()
    # 计数完全交给 WAA 自己的费用条状态机（暂停时帧号冻结在真实帧）；
    # 暂停/倍速/子弹时间按钮仍用 WAA 的识别。
    start_offset = 0
    # 自动检测只在轴里声明的档位之间切换；--cost-rate 调试时固定该值。
    rate_candidates = (declared if (declared and args.cost_rate is None)
                       else [rate])
    print("自动检测候选回费速度 = "
          + " / ".join(f"{x:g}%" for x in rate_candidates))
    counter = WAACounter(rate, neg_override=neg_override,
                     neg_speed=neg_speed, neg_declared=neg_declared,
                     rate_candidates=rate_candidates, require_hud=True)
    counter.csv_path = os.path.join(OUT, "waa_run.csv")
    if args.start_frame is not None:
        # 仅调试用：强制从关卡中途某个全局帧开始。
        counter.set_initial_frame(args.start_frame)
        print(f"initial frame forced: {args.start_frame}")
    elif not args.start_stage:
        counter.arm(0)
    counter_thread = FrameCounterThread(stream, counter)
    counter_thread.start()

    # 面板对象先建好；走 --start-stage 时会在编队界面、点“开始行动”
    # 之前打开它。没走 --start-stage（已经在局内）则仍由第一次 F6 打开。
    reporter = ExcelReporter(
        axis,
        {"base": rate, "declared": declared, "neg_speed": neg_speed},
        os.path.join(OUT, "run_reports"),
        enabled=not args.no_excel,
        template_path=os.path.join(ROOT, "axes", "WAA_sample_1-7.xlsx"),
        in_place=args.in_place,
    )

    if args.start_stage:
        # 先确认眼前是编队界面。上次失败就是因为停在“选择干员”界面，
        # (1655,756) 点到了干员卡上，然后白等 15 秒。
        start_icon_path = os.path.join(ROOT, "waa", "resources", "StartOperation.png")
        start_tpl = cv2.imread(start_icon_path, cv2.IMREAD_GRAYSCALE)
        if start_tpl is None:
            raise FileNotFoundError(start_icon_path)

        def _score_start(frame) -> float:
            if frame is None:
                return -1.0
            gfull = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            roi = gfull[500:1000, 1400:1900]
            best = -1.0
            for scale in (1.0, 0.9):
                tpl = start_tpl if scale == 1.0 else cv2.resize(
                    start_tpl, None, fx=scale, fy=scale,
                    interpolation=cv2.INTER_AREA)
                if tpl.shape[0] > roi.shape[0] or tpl.shape[1] > roi.shape[1]:
                    continue
                best = max(best, float(cv2.matchTemplate(
                    roi, tpl, cv2.TM_CCOEFF_NORMED).max()))
            return best

        frame0 = None
        deadline0 = time.time() + 3.0
        while time.time() < deadline0 and frame0 is None:
            frame0 = stream.latest()
            if frame0 is None:
                time.sleep(0.02)
        start_score = _score_start(frame0)
        print(f"编队界面检测 score={start_score:.3f}")
        if start_score < 0.55:
            if frame0 is not None:
                Image.fromarray(frame0[:, :, ::-1]).save(
                    os.path.join(OUT, "axis_wrong_screen.png"))
            print("当前不是编队界面（没检测到“开始行动”按钮），已停止；"
                  "请先回到编队界面再运行。")
            counter_thread.stop()
            counter.close_log()
            stream.stop()
            touch.close()
            return 2

        # 用户要求：在编队界面、点“开始行动”之前把 Excel 面板打开。
        # 这时战斗还没开始，开表不会跑掉帧；开完必须把游戏重新拉到前台，
        # 否则“开始行动”的鼠标点击会落到 Excel 上。
        if not args.no_excel:
            print("开始作战前打开 Excel 面板…")
            if not reporter.start(wait_ready=True):
                print("[excel] 面板没打开，已停止，不点击“开始行动”。"
                      "请先保存并关闭占用轴文件的 Excel。")
                reporter.stop()
                counter_thread.stop()
                counter.close_log()
                stream.stop()
                touch.close()
                return 4
            if not _focus_window(hwnd):
                print("WARNING: 打开 Excel 后游戏窗口没到前台")
        else:
            print("--no-excel：跳过 Excel 面板")

        left, top = win32gui.ClientToScreen(hwnd, (0, 0))
        _, _, cw, ch = win32gui.GetClientRect(hwnd)
        # 编队界面“开始行动”蓝色按钮中心（1920×1080 实测 1655,756）。
        sx = left + int(0.862 * cw)
        sy = top + int(0.700 * ch)
        icon_path = os.path.join(ROOT, "waa", "resources",
                                 "BattleOfficiallyBegin.png")
        icon = cv2.imread(icon_path, cv2.IMREAD_GRAYSCALE)
        if icon is None:
            raise FileNotFoundError(icon_path)

        def _score_hud(frame) -> float:
            if frame is None:
                return -1.0
            gray = cv2.resize(cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY),
                              (1280, 720))
            roi = gray[20:85, 1165:1240]
            return float(cv2.matchTemplate(
                roi, icon, cv2.TM_CCOEFF_NORMED).max())

        # 点一次开始行动，没进游戏（还在编队界面）就重试；如果已经在
        # 加载（不在编队界面）就只等 HUD，不重复点击。
        max_attempts = 3
        per_wait = 10.0
        detected = False
        for attempt in range(1, max_attempts + 1):
            if attempt > 1:
                time.sleep(0.5)
                sc = _score_start(stream.latest())
                print(f"编队界面重试检测 score={sc:.3f}")
                if sc < 0.55:
                    print("当前不在编队界面（可能在加载），不再重复点击")
                    break
            # 编队界面用真实鼠标点击；触摸在这个界面上不触发按钮。
            # Excel 可能又抢了前台，点之前再确认一次。
            if not _focus_window(hwnd):
                print("WARNING: 点击“开始行动”前游戏窗口没到前台")
            mouse_click(sx, sy)
            print(f"第 {attempt}/{max_attempts} 次点击“开始行动” "
                  f"({sx},{sy})，等待战斗 HUD…")
            deadline = time.time() + per_wait
            while time.time() < deadline:
                sc = _score_hud(stream.latest())
                if sc >= 0.65:
                    print(f"战斗 HUD 出现 score={sc:.3f}")
                    # 编队界面的 UI 可能被误读；进关后等真正的费用条再起锚。
                    counter.arm(0)
                    detected = True
                    break
                time.sleep(0.02)
            if detected:
                break
            print(f"第 {attempt} 次点击后 {per_wait:.0f} 秒未检测到战斗 HUD")

        if not detected:
            # 可能已经点了、只是加载慢；最后再等一次，不再点击。
            print("最后等待战斗 HUD（不再点击）…")
            deadline = time.time() + 10.0
            while time.time() < deadline:
                sc = _score_hud(stream.latest())
                if sc >= 0.65:
                    print(f"战斗 HUD 出现 score={sc:.3f}")
                    counter.arm(0)
                    detected = True
                    break
                time.sleep(0.02)

        if not detected:
            frame = stream.latest()
            if frame is not None:
                Image.fromarray(frame[:, :, ::-1]).save(
                    os.path.join(OUT, "axis_no_hud.png"))
            raise RuntimeError(
                f"点击“开始行动” {max_attempts} 次后仍未检测到战斗开始")
        # HUD 会先于费用条出现；等费用条真正画出来再开始执行，
        # 否则执行器会在计数器 idle 时就去暂停/过帧。
        bar_ready = False
        deadline2 = time.time() + 8.0
        while time.time() < deadline2:
            if counter.armed and counter.last is not None:
                bar_ready = True
                break
            time.sleep(0.05)
        print(f"费用条出现 ready={bar_ready} "
              f"phase={getattr(counter.last, 'phase', None)} "
              f"width={getattr(counter.last, 'width', None)}")
        if not bar_ready:
            raise RuntimeError("检测到战斗 HUD，但 8 秒内没等到费用条")

    # 开局 HUD 是渐入的：等暂停键/倍速键连续读出 3 帧再交给执行器，
    # 否则按钮状态不确定时执行器会把正在跑的游戏点停。
    deadline3 = time.time() + 3.0
    while time.time() < deadline3 and counter.readable_streak < 3:
        time.sleep(0.05)
    print(f"按钮识别 streak={counter.readable_streak}")

    # 开局：倍速按钮一出现就立即暂停（AFA 的 GameMonitor 也是等
    # 倍速按钮白色像素出现后马上发暂停键，而不是等 running 图标）。
    print("等待开局倍速按钮出现（出现就立即暂停）…")
    deadline4 = time.time() + 30.0
    speed_ready = False
    while time.time() < deadline4:
        last = counter.last
        if last is not None and getattr(last, "speed_drawn", False):
            speed_ready = True
            break
        time.sleep(0.02)
    if not speed_ready:
        print("WARNING: 30 秒内没等到倍速按钮，退回等待 running=True")
        while time.time() < deadline4 + 10.0:
            last = counter.last
            if last is not None and getattr(last, "running", None) is True:
                break
            time.sleep(0.05)
    print(f"开局倍速按钮 ready={speed_ready} "
          f"running={getattr(counter.last, 'running', None)} "
          f"speed_drawn={getattr(counter.last, 'speed_drawn', None)} "
          f"frame={counter.frame}")

    executor = Executor(hwnd, axis, stream, touch, counter, start_offset=start_offset)
    executor.select_pause_mode = args.pause_mode
    if args.speed != "auto":
        executor.speed_override = int(args.speed)
    executor.end_pause = end_pause
    executor.step_enabled = not args.no_step
    executor.pause_method = args.pause_method

    # 能暂停时立即停住：不再先等 0.3 秒、也不先切倍速。
    # 停住后再切 2X；若开始前暂停=n，等 0.1 秒恢复；
    # 若=y，游戏保持暂停，等用户按 F6。
    print("开局能暂停：立即暂停游戏…")
    executor.pause_immediate()
    time.sleep(0.3)
    last = counter.last
    print(
        f"waa anchor: frame={counter.frame} "
        f"phase={getattr(last, 'phase', None)} "
        f"width={getattr(last, 'width', None)} "
        f"cap={getattr(last, 'cap_diff', 0.0):.1f} "
        f"running={getattr(last, 'running', None)} "
        f"m={getattr(last, 'm', None)} "
        f"bullet={getattr(last, 'bullet', None)} "
        f"neg={getattr(last, 'neg', None)} "
        f"spec_err={counter.counter.spec_error or '-'}"
    )
    if args.speed == "auto" and default_2x:
        if not executor.ensure_speed(2):
            print("WARNING: 轴设置要求默认 2 倍速，但没能切换成功；"
                  "请确认右上角是 2X")
    if start_pause:
        executor.manual_pause.set()
        print("轴设置：开始前暂停；游戏保持暂停，按 F6 恢复")
    else:
        time.sleep(0.1)
        executor.resume()
        print("轴设置：开始前暂停=n；已暂停 0.1 秒，恢复游戏")
    print(
        f"select_pause_mode = {executor.select_pause_mode} "
        f"speed={args.speed} step_enabled={executor.step_enabled} "
        f"pause_method={executor.pause_method}"
    )
    # Excel 面板在 --start-stage 时已经提前打开（见上）；这里只绑定运行时。
    # 没走 --start-stage 时面板还没开，仍由第一次 F6 手动暂停时打开。
    reporter.bind_runtime(counter, executor)
    executor.reporter = reporter
    if not executor._focus_game():
        print("WARNING: 游戏窗口没到前台；请点一下游戏窗口")
    start_pause_hotkey(executor)
    if args.start_stage and not args.no_excel:
        print("F5 = 打轴：作战记录末尾记录当前帧；"
              "F6 = 手动暂停 / 恢复；Excel 面板已在开始作战前打开")
    elif args.no_excel:
        print("F6 = 手动暂停 / 恢复（Excel 面板已禁用；F5 打轴不可用）")
    else:
        print("F5 = 打轴：作战记录末尾记录当前帧；"
              "F6 = 手动暂停 / 恢复；第一次按 F6 暂停后会自动打开 Excel 面板")
    exit_code = 0
    try:
        try:
            executor.run(max_actions=args.max_actions,
                         start_action=args.start_action)
            reporter.finish(True)
        except FatalReloadError as e:
            executor.axis_running = False
            reporter.set_error(str(e))
            reporter.finish(False)
            print(f"[错误] 轴已停止：{e}")
            exit_code = 1
        except Exception as e:
            executor.axis_running = False
            reporter.set_error(f"轴运行失败：{e}")
            reporter.finish(False)
            print(f"[错误] 轴运行失败：{type(e).__name__}: {e}")
            traceback.print_exc()
            exit_code = 1

        # 用户要求：操作做完后程序不结束，继续实时显示当前帧和回费速度。
        # Excel 面板线程、计数器、抓屏都保持运行；F6 仍可暂停/恢复。
        print("轴操作已结束。程序保持运行，实时刷新当前帧和回费速度；"
              "按 Ctrl+C 退出。")
        while win32gui.IsWindow(hwnd):
            if executor.manual_pause.is_set():
                _post_run_manual_pause(executor, reporter)
            time.sleep(0.1)
        print("游戏窗口已关闭，程序退出。")
    except KeyboardInterrupt:
        print("收到 Ctrl+C，退出实时显示。")
    finally:
        reporter.stop()
        counter_thread.stop()
        counter.close_log()
        stream.stop()
        touch.close()
        try:
            Image.fromarray(grab_client(hwnd)[:, :, ::-1]).save(
                os.path.join(OUT, "axis_after.png"))
        except Exception:
            pass
    print("done")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
