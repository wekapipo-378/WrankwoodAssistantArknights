"""打轴执行器（全局帧版）。

时间基准是 WAA 自己的费用条状态机（暂停帧不补 +1），不再读费用数字、不做 OCR。

执行流程：
    等到目标帧 → 暂停 → 找干员头像 → 拖到格子 → 选朝向 / 技能 / 撤退。
目标帧由轴里的 `帧号` 列给出（从开始计数那一刻算 0）。
"""

from __future__ import annotations

import csv
import os
import threading
import time

import cv2
import numpy as np
import win32api
import win32gui

from .buttons import read_buttons
from .frame_counter import normalize_frame
from .waa_counter import WAACounter
from .axis import Axis, AxisAction, read_axis
from .capture_stream import CaptureStream
from .send_input import key_down, key_up
from .touch_input import TouchDevice
from .vision import locate_avatar, operator_cards, rightmost_card_info
from .excel_sort import (FormulaRewriteError, collect_formula_rewrites,
                         compute_row_map)


class FatalReloadError(RuntimeError):
    """F6 排序/公式重写失败：必须停住，不能继续打轴。"""


class StopRequested(RuntimeError):
    """用户输入 stop；从等待/手动暂停中安全退出，不前置游戏。"""


def foreground_is(hwnd: int) -> bool:
    """前台窗口是不是这个窗口（子窗口归一到顶层再比）。

    Unity 客户端有时把焦点放在子窗口上；直接 `GetForegroundWindow() == hwnd`
    会误判成“游戏不在前台”，于是从 Excel 切回来时白闪一次窗口还是失败。
    """
    try:
        fg = win32gui.GetForegroundWindow()
        if not fg:
            return False
        if fg == hwnd:
            return True
        return win32gui.GetAncestor(fg, 2) == hwnd
    except Exception:
        return False

PAUSE_RATIO = (0.945, 0.075)
# AFA 的暂停按钮左右两半：左半继续，右半暂停。
PAUSE_LEFT_RATIO = (0.9400, 0.0700)
PAUSE_RIGHT_RATIO = (0.9650, 0.0700)
# 暂停/运行、1X/2X、子弹时间都由 waa.core.buttons 自己识别（几何结构，不用模板）。
PAUSE_BUTTON_CENTER = (1200 / 1280, 50 / 720)
# MAA BattleOperRetreatJustClick 的 PC 端撤退按钮位置。
RETREAT_BUTTON_RATIO = (572 / 1280, 225 / 720)
# MAA BattleCancelSelection：点顶部中间取消当前选中，镜头随之复位。
CANCEL_SELECTION_RATIO = (640 / 1280, 10 / 720)
# prts-plus 进子弹时间用的最后一个待部署卡牌位置。
BULLET_OPER_RATIO = (0.95, 0.90)
# 进子弹统一点这个固定点：待部署卡行是右对齐的，最右那张一直落在这里。
# 用旗标去“找最右卡”会被选中态/动画骗到右数第二张，点下去把选中切走
# （2026-11-09 1020 帧的装置技能就是这么废的）。
DEPLOY_CARD_RATIO = (15 / 16, 9 / 10)
# 右上角 1X/2X 按钮中心（1920×1080）。
SPEED_BUTTON_RATIO = (0.860, 0.067)
DIRECTION_DX = 0.20
DIRECTION_DY = 0.20
# 部署拖拽：点卡后等 0.2s，拖到中间等 0.2s，10 步 × 20ms，
# 松手后等 0.2s。游戏在子弹时间下也需要时间处理卡片选中/落点
# （2026-11-09 按用户要求把两个 0.4s 都改成 0.2s）。
DEPLOY_DRAG_STEPS = 10
# 每步 20ms：比游戏帧（16.7ms）略长，保证每个中间点都被处理一次。
# 12ms 时相邻移动会被合并，落点容易取到中间值。
DEPLOY_DRAG_STEP = 0.02
CARD_CLICK_WAIT = 0.2
DRAG_MIDDLE_WAIT = 0.2
DROP_WAIT = 0.2
# 到达落点后、抬手前的停顿。拖拽步长 12ms < 游戏帧 16.7ms，
# 移动和抬起挨太近时游戏可能按上一个中间点判定落点。
DROP_HOLD = 0.08
# 基础帧步长：69905 / 2^21 = 0.0333333015s（不是 1/30）。
# 子弹时间固定 0.2×，与 1X/2X 无关；一步 = 5 个基础帧。
BULLET_STEP_PERIOD = 5 * 69905 / 2**21
# 右半“暂停”点完后，等游戏确认暂停的时间。实测减半会在
# 29->31、597->599 这类位置过冲 2 帧，必须保持完整一拍。
BULLET_PAUSE_SETTLE = BULLET_STEP_PERIOD
# 过帧等待 = period - 这个值（秒）。
# 实测（瞬时 tap、每档 8 步、物理帧数用费用条宽度核对）：
#   0.000 → 8/8 一步一帧、0 过冲、0 重试、均 397ms（最好）
#   0.010 → 8/8、0 重试
#   0.020 → 7/8、1 次 +2、2 重试
# 两边都是瞬时 tap 时继续/暂停延迟对称、互相抵消，等满一整拍才对。
PAUSE_LATENCY_MARGIN = 0.0
# 「刚确认过暂停」的信任窗口（秒）。ESC 暂停后视觉样本滞后 40~100ms，
# 这段时间内不要再拿图标读数去点暂停，否则会把刚停住的游戏点开。
PAUSE_TRUST_S = 1.5
# 暂停提前量（lead）的单位统一用秒：暂停延迟是墙钟时间，跟倍速无关。
# 下面两个常量是旧「安全下限」（实测延迟 + 50ms）用的，2026-11-09 按用户要求
# 停用：严格暂停的提前量现在直接取轴里填的秒，填 0ms 就是 0。
# 保留定义只是为了不在别处留下悬空引用；_strict_pause_lead_s 不再使用它们。
PAUSE_LEAD_MARGIN_S = 0.05
DEFAULT_PAUSE_LATENCY_S = 0.14
# 一次步进最多尝试 5 次（0 帧就重试）。
STEP_ATTEMPTS = 5
# 进/出子弹、选人、E/Q 的确认上限。
BULLET_CONFIRM = 0.3
# 退出子弹：点击后最多 20ms 就要判定。
EXIT_CONFIRM = 0.02
# 但 bullet 状态来自视觉线程（WGC 60Hz + 计数线程逐帧消费），通常滞后
# 1~3 帧。这 20ms 内如果连一帧新样本都没来，读到的还是点击前的旧状态，
# 不能据此认定“没退出”——那会接着点第二次（点卡 = 重新进子弹）。
# 这种情况最多再等 40ms，看到新样本立刻下结论。
EXIT_CONFIRM_GRACE = 0.04
# 认为“状态是新的”所需的最小样本推进量：约 2 个 WGC 帧（60Hz）。
# 点击后第一帧新样本可能还是在游戏处理点击之前抓的，不能拿它下结论。
EXIT_FRESH_MIN = 0.03
# E/Q 之后等子弹时间消失的确认时间。技能真的放出来时一两帧就消失
# （30~50ms），0.3s 只是让没充能的技能白等 3 次；实测帧 240 为此停了 1.7s。
KEY_CONFIRM = 0.15
# 计数器落点稳定上限。
SETTLE_TIMEOUT = 0.5
# 单次 run 里帧号倒退超过这么多，就认定已经离开战斗/重开，直接停轴。
FRAME_ROLLBACK_LIMIT = 120
SKILL_KEY = 0x45    # E
RETREAT_KEY = 0x51  # Q
VK_ESCAPE = 0x1B
VK_SPACE = 0x20
# 连续技能/撤退相隔这么多帧以内，放进同一个暂停窗口连续执行。
# 相差很近的操作绝不恢复暂停，直接在当前暂停里一帧一帧走到下一个动作。
BATCH_WINDOW = 10
# 技能/撤退提前这么多帧暂停：先选中干员进入子弹时间，再逐帧走到目标帧。
PRE_STEP_FRAMES = 2
# E/Q 之后等镜头复位再继续。
ACTION_SETTLE = 0.5
# 选中重试次数。真实轴会遇到技能没放出、Q 没反应的情况，不能按完就当成功。
SELECT_RETRIES = 3
# 前置轻点（direct）之后、三连点之前的等待。这一步只是“唤醒”客户端，
# 不需要半秒；实测 500ms 让每个技能动作多花 0.5s（帧 240 因此停了 1.7s）。
W_SETTLE = 0.2
# 三连点后只等 ClickDelay 就按键（对齐 AFA 的 ClickDelay 默认 50ms）。
CLICK_DELAY = 0.05
W_TAP_GAP = 0.01


class Executor:
    def __init__(self, hwnd: int, axis: Axis, stream: CaptureStream,
                 touch: TouchDevice, counter: WAACounter, start_offset: int = 0):
        self.hwnd = hwnd
        self.axis = axis
        self.stream = stream
        self.touch = touch
        self.counter = counter
        # 通常为 0：WAA 自己会从当前费用条相位起锚，不能再重复加偏移。
        self.start_offset = int(start_offset)
        self.paused = True
        # 上一次「确认已暂停」的时刻（monotonic）。ESC 暂停后视觉样本会滞后
        # 40~100ms，pause_immediate 若立刻读图标会读到旧的“运行中”并把游戏
        # 点开；这个时间戳给出一段直接信任的窗口。
        self._paused_at = 0.0
        # AFA 三连点：左半继续 → 点干员 → 右半暂停（按住 30ms）；
        # 用户要求选中干员一律用这个动作。
        self.select_pause_mode = "hold"
        self._esc_held = False
        # None = 从 WAA 读倍速；1/2 = 手动覆盖。
        self.speed_override = None
        # 测试用：关闭单帧步进，避免步进点卡牌干扰撤退。
        self.step_enabled = True
        # 用户手动用 ESC 暂停/再 ESC 恢复；默认仍用触控按钮。
        self.pause_method = "touch"
        # 调试输出开关：默认关。开（verbose=True）时才会打印那些逐帧/坐标/
        # 重试的细节；关闭时屏幕只保留动作结果、暂停落点、WARNING/ERROR。
        self.verbose = False
        # 每次暂停的 (trigger, landing, latency, lead)，用于找稳妥提前量。
        self.pause_samples: list[tuple[int, int, int, int, float]] = []
        # 每次操作的 trigger 帧 / 实际暂停帧，运行结束写入 out/pause_log.csv。
        self.pause_log: list[dict] = []
        self.last_pause_trigger: int | None = None
        self.last_pause_landing: int | None = None
        # F6 手动暂停：脚本停住游戏、帧号冻结；用户切后台看 Excel，再按 F6 恢复。
        self.manual_pause = threading.Event()
        # 部署的拖拽/方向阶段由主线程自己按顺序处理 F6（回拖、取消子弹、
        # 再暂停），热键线程在这段时间不要抢先点暂停。
        self.deploy_active = False
        # 技能/撤退最近一次“点最右待部署卡进子弹”的屏幕坐标；到目标帧
        # 后再点它一次退出子弹时间。
        self._last_deploy_card_screen = None
        # 本次动作进入子弹时选中的最右卡；如果目标正好是这张卡，
        # 部署时不要再点一次（再点会把它取消）。
        self._bullet_entry_card_screen = None
        # 当前子弹时间里选中的单位名（地图页“当前状态”显示）。
        self.selected_oper: str | None = None
        # 轴里的“暂停”操作：已停住，等待用户手动恢复中。
        self.manual_resume_wait = False
        # 轴设置“结束后暂停”：最后一个动作完成后不要 resume。
        self.end_pause = False
        # Excel 运行面板（可为 None）；由 run_axis 注入。
        self.reporter = None
        self.current_action = None
        # 当前/之后还没执行的动作在 axis.actions 里的起始下标；F6 重载只
        # 更新这个下标之后的行，已执行动作的结果不改。
        self.pending_from = 0
        # 运行时部署栈：记录实际落点，进子弹的备用入口和撤退会用到。
        self.deployed_stack: list[dict] = []
        # 运行信息页“当前状态”显示的阶段。
        self.phase = "准备中"
        # 轴是否正在执行；F5 只在轴运行期间触发“打轴时暂停”。
        self.axis_running = False
        # F5 的独立写入目标；None=写主轴的 ExcelReporter。
        self.mark_writer = None
        # write 文件选过但打不开/不存在：F5 报错，不回退到主轴。
        self.mark_writer_error = ""
        # read 文件对话框/读取期间暂时拒绝 F5。
        self.mark_suspended = False
        # 控制台 stop 请求。
        self.stop_requested = threading.Event()

        from .capture import client_rect
        self.client_left, self.client_top, self.cw, self.ch = client_rect(hwnd)
        self._last_running: bool | None = None
        self._last_icon_score = 0.0

    # ---- 坐标 ----
    def screen(self, rx: float, ry: float) -> tuple[int, int]:
        return self.client_left + int(rx * self.cw), self.client_top + int(ry * self.ch)

    # ---- 帧 ----
    def current_frame(self) -> int:
        return self.counter.frame + self.start_offset

    def current_frame_hi(self) -> int:
        """当前帧的上界：费用条多帧同宽时，真实帧可能在 frame~hi 之间。"""
        return self.counter.frame_hi + self.start_offset

    # ---- 输出分级 ----
    def _dbg(self, msg: str) -> None:
        """调试噪音：默认不打印，只有 verbose 开启（debug/verbose 命令）才显示。

        用于逐帧坐标、选卡重试、子弹进出、步进细节这类排查用的输出；
        动作结果、暂停落点、WARNING/ERROR 走普通 print，始终显示。
        """
        if getattr(self, "verbose", False):
            print(msg)

    # ---- 控制台接口 ----
    def set_offset(self, delta: int | None = None,
                   value: int | None = None) -> int:
        if value is not None:
            self.start_offset = int(value)
        elif delta is not None:
            self.start_offset += int(delta)
        print(f"  offset = {self.start_offset}")
        return self.start_offset

    def set_frame(self, frame: int) -> int:
        self.start_offset = int(frame) - int(self.counter.frame)
        print(f"  frame = {frame}（offset = {self.start_offset}）")
        return self.start_offset

    def request_stop(self) -> None:
        self.stop_requested.set()
        print("  收到 stop：安全退出，不切回游戏窗口")

    def _raise_if_stop(self) -> None:
        ev = getattr(self, "stop_requested", None)
        if ev is not None and ev.is_set():
            raise StopRequested("用户停止")

    # ---- 暂停 ----
    def _tap_esc(self) -> None:
        key_down(VK_ESCAPE)
        time.sleep(0.05)
        key_up(VK_ESCAPE)

    def _tap_space(self) -> None:
        """按一次 Space（AFA 的 pauseBattle 键）。AFA game_monitor 开局暂停就是
        倍速按钮白色像素一出现就 `SendDown(pauseBattle=Space)` + USleep(50) +
        SendUp，只发一次、不确认。"""
        key_down(VK_SPACE)
        time.sleep(0.05)
        key_up(VK_SPACE)

    def _click_pause(self) -> None:
        if self.pause_method == "esc":
            self._tap_esc()
            return
        # 以前这里按住 50ms，注释写“10ms 太短，游戏有时不认继续”。
        # 2026-11-08 管理员权限实测（out/_tap_probe*.py）：按钮中心用
        # 瞬时 tap 双向切换 7/7 成功，暂停 88ms、继续 91ms 就能读到翻面，
        # 比按住 50ms（总耗时 100~130ms）更快。按住纯属白等。
        self.touch.tap(*self.screen(*PAUSE_RATIO))

    def _advancing(self, wait: float = 0.12) -> bool:
        """看 WAA 帧号是否真的在涨，而不是相信自己记的 paused 布尔值。"""
        before = self.counter.frame
        time.sleep(wait)
        return self.counter.frame > before

    def _wait_pause_state(self, want_running: bool,
                          timeout: float = 0.15) -> bool:
        """轮询右上角图标，翻面就立刻返回。

        以前点完固定 `sleep(0.15)`：图标其实 1~2ms 就变了，每次暂停/继续
        都白等 150ms。这里改成边读边等，最快一拍就返回，超时上限不变。
        """
        deadline = time.perf_counter() + timeout
        while True:
            running = self._read_running()
            if running is not None and running == want_running:
                return True
            if time.perf_counter() >= deadline:
                return False
            time.sleep(0.008)

    def _read_running_stable(self, gap: float = 0.02) -> bool | None:
        """连读两次一致才算数。

        intro 阶段暂停键是淡入的，单次读数会把残留图形认成“单个三角=暂停”
        （2026-11-09 六轮开局暂停全都栽在这上面）。
        """
        first = self._read_running()
        if first is None:
            return None
        time.sleep(gap)
        second = self._read_running()
        return first if first == second else None

    def pause_by_key(self, wait_armed: float = 10.0,
                     wait_visible: float = 1.6) -> bool:
        """开局用：倍速按钮的白色像素一出现就发一次 Space，**不确认**。
        这是 AFA 自动开局暂停的精确复刻。

        AFA 的 game_monitor（game_monitor.ahk `_PauseWaitTick`）：`SafePixelSearch`
        检测倍速按钮区域（0.845~0.8807, 0.0713~0.087，即 1920×1080 下 x1622~1691,
        y77~94）的**纯白像素 0xffffff（容差 10）**，一命中就
        `SendDown(pauseBattle=Space)` + USleep(50) + SendUp，**只发一次、不确认**。

        关键：必须用「纯白像素」这个精确信号，不能用 WAA 的 speed_drawn（相对
        亮度阈值）。后者在加载/intro 阶段（画面偏暗、只有相对亮块）就误判成
        True，导致 Space 发早被吞（2026-11-09 14:47 轮：speed_drawn t=12.96 就=1、
        但游戏 t=13.76 才开始回费，Space 白按）。纯白像素是倍速按钮文字真正
        画出来的时刻，比费用条 armed 早、又不会在加载阶段误判。

        风险与兜底：Space 万一仍被吞，交给第一个 pause_at 兜底（等同旧版开局
        暂停失败）。这是用户明确选择“追平 AFA、接受小概率漏帧”。
        """
        from .buttons import speed_white_pixel
        deadline = time.perf_counter() + wait_armed
        while time.perf_counter() < deadline:
            frame = self.stream.latest()
            if frame is None:
                time.sleep(0.02)
                continue
            gray = normalize_frame(cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY))
            if speed_white_pixel(gray):
                break
            time.sleep(0.02)
        else:
            print(f"  开局暂停：{wait_armed:g}s 内没等到倍速按钮白像素，"
                  f"放弃开局暂停（不硬按，交给第一个 pause_at）")
            return False
        # 倍速按钮白像素刚出现：发一次 Space（AFA 的 pauseBattle），不确认。
        self._tap_space()
        self.paused = True
        self._paused_at = time.monotonic()
        print("  开局暂停：倍速按钮白像素出现，已发 Space（不确认，照搬 AFA）")
        return True

    def pause_immediate(self) -> None:
        """开局用：不等确认，直接点一次暂停；已经在暂停就什么都不做。

        **不能拿一次滞后读数去点**：ESC 暂停后视觉样本还落后 40~100ms，
        这时 is_running_icon() 读到的是暂停前的“运行中”，一点就把刚停住的
        游戏点开。2026-11-09 06:16 实测就是这条链路：

            t=17.29 frame=2 running=0   ← ESC 停住
            t=18.07 frame=2 running=1   ← pause_immediate 又点了一下
            t=18.16 frame=7 running=0   ← pause() 重试才停住

        所以程序自己记“刚确认过暂停”的时刻，1.5s 内直接信任；
        过了窗口也要求连读两次一致才认为在运行。
        """
        if self.paused:
            if time.monotonic() - self._paused_at < PAUSE_TRUST_S:
                return
            running = self._read_running_stable()
            if running is not True:
                # False = 确实停着；None = 读不到（intro 淡入）——都不该点。
                self.paused = True
                self._paused_at = time.monotonic()
                return
        elif not self.is_running_icon():
            self.paused = True
            self._paused_at = time.monotonic()
            return
        self._click_pause()
        if self._wait_pause_state(False, 0.2):
            self.paused = True
            self._paused_at = time.monotonic()
            return
        # 第一次没停住再走常规重试，避免误判成暂停。
        self.pause(assume_running=True)

    def pause(self, assume_running: bool = False, retries: int = 4) -> None:
        """暂停：以右上角按钮为准。

        开局 intro 阶段费用条帧号本来就不涨（还没开始回费），
        旧版用“帧号是否在涨”判断暂停，会把 intro 误当成暂停并乱点。
        """
        for _ in range(retries):
            running = self._read_running()
            if running is None:
                time.sleep(0.1)
                continue
            if not running:
                self.paused = True
                self._paused_at = time.monotonic()
                return
            self._click_pause()
            if self._wait_pause_state(False, 0.15):
                self.paused = True
                self._paused_at = time.monotonic()
                return
        # 按钮一直读不到才退回看帧号
        if not self._advancing():
            self.paused = True
            self._paused_at = time.monotonic()
            return
        print("  WARNING: pause 失败，游戏仍在运行")

    def resume(self, assume_paused: bool = False, retries: int = 4) -> None:
        """恢复运行：同样以按钮为准，读不到就不点。"""
        for _ in range(retries):
            running = self._read_running()
            if running is None:
                time.sleep(0.1)
                continue
            if running:
                self.paused = False
                self._paused_at = 0.0
                return
            self._click_pause()
            if self._wait_pause_state(True, 0.15):
                self.paused = False
                self._paused_at = 0.0
                return
        if self._advancing():
            self.paused = False
            self._paused_at = 0.0
            return
        print("  WARNING: resume 失败，游戏仍在暂停")

    def _speed(self) -> float:
        """当前倍速：优先手动覆盖，否则从 WAA 读数取 1X/2X。"""
        if self.speed_override in (1, 2):
            return float(self.speed_override)
        m = getattr(self.counter.last, "m", 1.0) if self.counter.last is not None else 1.0
        try:
            m = float(m)
        except (TypeError, ValueError):
            m = 1.0
        return 2.0 if m >= 1.5 else 1.0

    @staticmethod
    def _parse_lead_seconds(raw, speed: float) -> float | None:
        """解析轴里填的提前量，统一返回**秒**；认不出来回 None。

        只认带时间单位的写法（不区分大小写）：
            120ms / 120毫秒
            0.12s / 0.12秒
        裸数字（7）、7帧、7% 一律不认——避免“7 到底是帧还是秒”的歧义，
        认不出来就回落到默认值（2X 5 帧 / 1X 4 帧）。
        """
        if raw is None:
            return None
        text = str(raw).strip().lower().replace(" ", "")
        if not text:
            return None
        for unit in ("毫秒", "ms"):
            if text.endswith(unit):
                try:
                    return float(text[:-len(unit)]) / 1000.0
                except ValueError:
                    return None
        for unit in ("秒", "s"):
            if text.endswith(unit):
                try:
                    return float(text[:-len(unit)])
                except ValueError:
                    return None
        return None

    def _default_lead_s(self) -> float:
        """轴设置里的提前量，返回秒（1X 读 M7、2X 读 M8；空白默认 4/5 帧）。"""
        speed = self._speed()
        key = "2x操作提前量" if speed >= 1.5 else "1x操作提前量"
        default_frames = 5 if speed >= 1.5 else 4
        settings = getattr(getattr(self, "axis", None), "settings", {}) or {}
        raw = settings.get(key)
        seconds = self._parse_lead_seconds(raw, speed)
        if seconds is None:
            text = "" if raw is None else str(raw).strip()
            if text and not getattr(self, "_lead_warned", False):
                self._lead_warned = True
                print(f"  WARNING: 「{key}」填的 {text!r} 认不出来（只认 120ms / 0.12s）；"
                      f"已按默认 {default_frames} 帧处理")
            seconds = default_frames / (30.0 * speed)
        return max(0.0, seconds)

    def _default_lead(self) -> int:
        """提前量的帧数写法（旧接口）：由秒按当前倍速换算。"""
        return max(1, int(round(self._default_lead_s() * 30.0 * self._speed())))

    def _strict_pause_lead_s(self) -> float:
        """严格暂停的提前量，单位**秒**：完全等于轴里填的值。

        旧实现用 max(轴值, 最近实测暂停延迟 + 50ms)，保证怎么填都不晚帧，
        代价是用户填 0ms / 小值时被下限吞掉，看起来「调了没反应」。
        2026-11-09 按用户要求先去掉下限、再去掉 0.4s 上限：现在既不设下限
        也不设上限，填多少就是多少（负数按 0 处理）。

        提前量小于实测暂停延迟时 landing 会超过 target；_run_actions 的
        严格不晚帧检查会照常打印 WARNING 并放弃该动作，这里不吞失败、
        也不放宽判据。填超大值会提前很多帧暂停、白走同样多的子弹步，
        由轴作者自己负责。
        """
        return max(0.0, self._default_lead_s())

    def _frame_period(self) -> float:
        return 1.0 / (30.0 * self._speed())

    @staticmethod
    def _sleep_precise(seconds: float) -> None:
        end = time.perf_counter() + max(0.0, seconds)
        while time.perf_counter() < end:
            pass

    def _read_running(self) -> bool | None:
        """读右上角暂停键：True=运行，False=暂停，None=还没画稳/读不到。"""
        frame = self.stream.latest()
        if frame is None:
            return None
        gray = normalize_frame(cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY))
        st = read_buttons(gray)
        self._last_icon_score = 1.0 if st.pause_drawn else 0.0
        if st.running is not None:
            self._last_running = st.running
        return st.running

    def is_running_icon(self) -> bool:
        """WAA 自己识别：双竖条=运行，单三角=暂停；读不到时沿用上一次。"""
        running = self._read_running()
        if running is None:
            # 开局按钮还没画稳：按“运行中”处理，宁可不点，也不要把游戏点停
            return True if self._last_running is None else self._last_running
        return running

    def is_paused(self) -> bool:
        return not self.is_running_icon()

    def ensure_speed(self, target: int = 2, retries: int = 4) -> bool:
        """把右上角速度切到 target（1/2），并确认 WAA 读到对应倍速。"""
        target = int(target)
        for attempt in range(retries):
            m = getattr(self.counter.last, "m", None) if self.counter.last is not None else None
            try:
                m = float(m) if m is not None else None
            except (TypeError, ValueError):
                m = None
            if m is not None and abs(m - target) < 0.5:
                self.speed_override = target
                self._dbg(f"  speed already {target}X")
                return True
            self.touch.click(*self.screen(*SPEED_BUTTON_RATIO), hold_ms=30)
            time.sleep(0.4)
            after = getattr(self.counter.last, "m", None)
            self._dbg(f"  ensure speed click {attempt + 1}: "
                      f"before={m} after={after} target={target}")
        m = getattr(self.counter.last, "m", None) if self.counter.last is not None else None
        try:
            ok = m is not None and abs(float(m) - target) < 0.5
        except (TypeError, ValueError):
            ok = False
        if ok:
            self.speed_override = target
        else:
            print(f"  WARNING: 速度没切到 {target}X（当前 {m}）")
        return ok

    def _settled_frame(self, timeout: float = SETTLE_TIMEOUT) -> int:
        """暂停后等计数线程把在途帧消费完，再读冻结的真实帧。

        新计数器在暂停时直接停在被暂停的那一帧，不需要再 +1；
        但要等队列里的旧帧走完，否则会读到稍早的帧号。
        """
        last = self.current_frame()
        stable = 0
        deadline = time.perf_counter() + timeout
        while time.perf_counter() < deadline:
            time.sleep(0.03)
            now = self.current_frame()
            if now == last:
                stable += 1
                if stable >= 2:      # 连续约 60ms 不变 = 停住了
                    break
            else:
                last = now
                stable = 0
        return self.current_frame()

    def _frame_now(self) -> int:
        """当前真实帧：停住（暂停/子弹）时先等在途帧走完再读。"""
        if self.is_paused() or self.bullet_state() is True:
            return self._settled_frame()
        return self.current_frame()

    # ---- 选中 / 按键确认（针对技能没放出、Q 没反应） ----
    def bullet_state(self) -> bool | None:
        """返回当前是否子弹时间；None=按钮还没读稳。"""
        last = self.counter.last
        if last is None:
            return None
        try:
            return bool(last.bullet)
        except Exception:
            return None

    def _wait_bullet(self, want: bool, timeout: float = 0.6) -> bool:
        deadline = time.perf_counter() + timeout
        state: bool | None = None
        while time.perf_counter() < deadline:
            state = self.bullet_state()
            if state is not None and state == want:
                return True
            time.sleep(0.03)
        return state == want

    def _rightmost_card_screen(self) -> tuple[tuple[int, int] | None, dict | None]:
        """当前最右卡的头像中心（屏幕坐标）+ 卡片信息。"""
        try:
            frame = self.stream.latest()
            card = rightmost_card_info(frame) if frame is not None else None
        except Exception as exc:
            print(f"  WARNING: 旗标找卡失败，改用固定位置: {exc}")
            return None, None
        if card is None:
            return None, None
        ax, ay = card["avatar_center"]
        pt = self.screen(ax / 1280.0, ay / 720.0)
        return pt, card

    def _click_fixed_deploy_card(self) -> bool:
        """点固定的待部署卡位置进子弹（不识别、不挑卡）。

        待部署卡行是右对齐的，最右那张一直落在这个固定比例上，
        所以没必要去找“最右卡”。找卡反而会被选中态/动画骗到右数第二张。
        """
        pt = self.screen(*DEPLOY_CARD_RATIO)
        self._dbg(f"  deploy card 固定点 ({pt[0]},{pt[1]})")
        self._last_deploy_card_screen = pt
        self.touch.click(*pt, hold_ms=30)
        return True

    def _click_rightmost_deploy_card(self) -> bool:
        """旗标找最右待部署卡并点击（固定点失败时才用）。"""
        screen_pt, card = self._rightmost_card_screen()
        if screen_pt is None:
            screen_pt = self.screen(*BULLET_OPER_RATIO)
            self._dbg(f"  rightmost card fallback fixed "
                      f"({BULLET_OPER_RATIO[0]:.3f},{BULLET_OPER_RATIO[1]:.3f})")
        else:
            ax, ay = card["avatar_center"]
            self._dbg(f"  rightmost card avatar ({ax / 1280.0:.3f},"
                      f"{ay / 720.0:.3f})")
        self._last_deploy_card_screen = screen_pt
        self.touch.click(*screen_pt, hold_ms=30)
        return True

    def _enter_bullet(self) -> bool:
        """兜底：点最右待部署卡进子弹时间。"""
        if not self.is_paused():
            self.pause()
        self._click_fixed_deploy_card()
        time.sleep(0.25)
        ok = self._wait_bullet(True, 0.5)
        self._dbg(f"  enter bullet fallback frame={self.current_frame()} "
                  f"bullet={self.bullet_state()} ok={ok}")
        if ok:
            self._bullet_entry_card_screen = self._last_deploy_card_screen
        return ok

    def _enter_bullet_by_card(self) -> bool:
        """技能/撤退第二步：点最右待部署卡进入子弹时间并确认。"""
        self._click_fixed_deploy_card()
        time.sleep(0.25)
        ok = self._wait_bullet(True, 0.5)
        self._dbg(f"  enter bullet by card: bullet={self.bullet_state()} "
                  f"frame={self.current_frame()} ok={ok}")
        if ok:
            self._bullet_entry_card_screen = self._last_deploy_card_screen
        return ok

    def _select_operator(self, action: AxisAction,
                         target_frame: int | None = None) -> str:
        """选中干员并确认出现 >>>（bullet）；失败就换方式重试。

        target_frame 用于区间技能：区间内当前帧可能不是 action.frame，
        但仍应允许重试，不能按 action.frame 判“已过帧”。
        """
        if action.view_front is None:
            raise ValueError(f"技能/撤退没有位置: {action}")
        target = int(action.frame if target_frame is None else target_frame)

        if not self.is_paused():
            self.pause()

        # 用户确认：暂停中单击干员选不中，必须“继续→选中→暂停”三连点。
        # 但这一下前置轻点不能删：23:35 那轮把它去掉后，5 次选人全是
        # bullet=False（tap3 显示三连点本身没跨帧）。保留它当“唤醒点击”，
        # 不指望它选中，真正选中仍靠后面的 hold 三连点。
        modes: list[str] = []
        if self.current_frame() >= target - 1:
            modes.append("direct")
        modes.extend([self.select_pause_mode, "hold", "tap"])
        deduped: list[str] = []
        for m in modes:
            if m not in deduped:
                deduped.append(m)
        modes = deduped

        for attempt in range(SELECT_RETRIES):
            for mode in modes:
                self._dbg(f"  select try={attempt + 1} mode={mode} "
                          f"frame={self.current_frame()} bullet={self.bullet_state()}")
                self.select_while_paused(*action.view_front, pause_mode=mode)
                ok = self._wait_bullet(True, 0.5)
                if not ok and not self.is_paused():
                    self.pause()
                # 三连点后计数器可能还有在途帧，直接读会把同一帧的选人
                # 误判成 target+1；统一读冻结帧再判断。
                cur = self._frame_now()
                self._dbg(f"  select verify mode={mode} bullet={self.bullet_state()} "
                          f"ok={ok} frame={cur}")
                if ok:
                    if cur > target:
                        raise RuntimeError(
                            f"选人时已过目标帧：目标 {target}，当前 {cur}")
                    self.selected_oper = getattr(action, "oper", None)
                    return mode
                if cur > target:
                    raise RuntimeError(
                        f"选人重试时已过目标帧：目标 {target}，"
                        f"当前 {cur}")
            if attempt == 0:
                self._enter_bullet()

        raise RuntimeError(
            f"选中干员失败（没有出现子弹时间 >>>）：{action}")

    def _click_exit_bullet_point(self) -> None:
        """点右上角往左 1/16 宽、往下 1/2 高的点，用来退出子弹时间。"""
        x = self.client_left + self.cw - self.cw // 16
        y = self.client_top + self.ch // 2
        self._dbg(f"  点击退出子弹点 ({x},{y})")
        self.touch.click(x, y, hold_ms=30)

    def _send_action_verified(self, action: AxisAction, vk: int,
                              label: str, kind: str) -> bool:
        """按 E/Q 并确认离开子弹时间。

        第一次失败后，同一帧最多再重试 3 次；仍失败则点固定退出点，
        强制退出子弹时间。动作本身记为失败，但不会把 bullet 留到下一动。
        """
        if self.bullet_state() is not True:
            self._select_operator(action)
        for attempt in range(3):          # 1 次 + 最多 2 次重试
            self._send_action_key(vk, label)
            ok = self._wait_bullet(False, KEY_CONFIRM)
            self._dbg(f"  verify {label} #{attempt + 1}: "
                      f"bullet={self.bullet_state()} "
                      f"frame={self.current_frame()} ok={ok}")
            if ok:
                return True
            if attempt < 2:
                self._dbg(f"  {kind}仍在子弹时间，重试 E/Q "
                          f"({attempt + 1}/2)")
        # 四次都没退出：点固定位置强制退出，避免后面的动作全错。
        self._click_exit_bullet_point()
        time.sleep(0.2)
        still = self._wait_bullet(False, 0.6)
        print(f"  WARNING: {kind}没有生效，已尝试退出子弹点；"
              f"bullet={self.bullet_state()} exited={still}")
        return False

    def _set_pause(self, paused: bool) -> None:
        """同一个暂停按钮切换，用图标确认状态，失败就再点。

        读不到按钮（开局 HUD 还没画稳）时 **不点**，只等；
        否则会把正在跑的游戏点停，这就是上一局开局失败的根因。
        """
        for _ in range(12):
            running = self._read_running()
            if running is None:
                time.sleep(0.15)
                continue
            if (not running) == paused:
                return
            # 左半只继续、右半只暂停；1920 下按钮中心可能偏到左边缘，
            # 用左右半比点中心稳定。
            ratio = PAUSE_RIGHT_RATIO if paused else PAUSE_LEFT_RATIO
            self.touch.click(*self.screen(*ratio), hold_ms=50)
            self._wait_pause_state(not paused, 0.2)
        print(
            f"  WARNING: 暂停状态切换失败 icon_score="
            f"{getattr(self, '_last_icon_score', None)}"
        )

    def force_pause(self) -> None:
        """兼容旧调用：确保暂停，但不会把暂停切成运行。"""
        self._set_pause(True)

    # ---- F5 打轴：记录当前帧 ----
    def _axis_setting_bool(self, key: str, default: bool) -> bool:
        """读轴设置里的布尔值；空白/无法识别时用默认值。"""
        settings = getattr(getattr(self, "axis", None), "settings", {}) or {}
        v = settings.get(key)
        if v is None or str(v).strip() == "":
            return default
        if isinstance(v, bool):
            return v
        s = str(v).strip().lower()
        if s in ("true", "1", "yes", "y", "是", "开", "on"):
            return True
        if s in ("false", "0", "no", "n", "否", "关", "off"):
            return False
        return default

    def _pause_after_mark(self) -> None:
        """F5 打轴后：按 F6 的暂停流程停住游戏，并前置 Excel。

        对齐 F6：热键线程只登记 manual_pause，**不直接点暂停**。主线程会在
        下一个安全点进入 `_manual_hold`，它已经实现「先暂停（is_running_icon
        判断，已暂停就不再按）→ 打开/前置 Excel → 等 F6 恢复」的完整流程。

        以前 F5 在这里直接 pause() + 单独前置 Excel，跟 _manual_hold 的暂停
        逻辑重复，而且热键线程直接做触摸注入不可靠（游戏窗口不在前台时注入
        被 UIPI 丢弃，2026-11-09 15:05 轮就是 pause 失败、游戏仍在运行）。
        """
        if self.manual_pause.is_set():
            # 已经在手动暂停中（_manual_hold 正在等 F6）：F5 只负责记录帧，
            # 暂停/前置 Excel 都由 _manual_hold 处理，这里不再重复动作。
            return
        if self.deploy_active:
            print("  F5：部署中，等部署流程到安全点再暂停")
            self.manual_pause.set()
            return
        # 正常运行：只登记，主线程安全点会走 _manual_hold（先暂停再开 Excel）。
        print("  F5：打轴暂停请求已登记，等主线程在安全点停住游戏")
        self.manual_pause.set()

    def mark_frame(self) -> int | None:
        """在作战记录末尾加一行，帧号1 填当前帧。

        设置“打轴时暂停”为 TRUE 时，加完行后按 F6 的流程暂停并前置 Excel。
        """
        frame = self.current_frame()
        if self.mark_suspended:
            print("  F5 打轴：正在读取轴文件，暂不可用")
            return None
        if self.mark_writer_error:
            print(f"  F5 打轴：{self.mark_writer_error}")
            return None
        writer = self.mark_writer
        if writer is not None:
            try:
                row = writer.add_axis_row(frame)
            except Exception as exc:
                print(f"  F5 打轴：写入 write 文件失败：{exc}")
                return None
            if row is None:
                print("  F5 打轴：write 文件写入失败")
                return None
            print(f"  F5 打轴：write 第 {row} 行，帧号1={frame}")
            if self.axis_running and self._axis_setting_bool(
                    "打轴时暂停", True):
                self._pause_after_mark()
            return int(row)
        rep = self.reporter
        if rep is None or not hasattr(rep, "add_axis_row"):
            print("  F5 打轴：还没有 read 轴文件，也没有 write 文件")
            return None
        row = rep.add_axis_row(frame)
        if row is None:
            print("  F5 打轴：写入 Excel 失败")
            return None
        self.axis.actions.append(AxisAction(
            frame=frame, kind="无操作", excel_row=int(row), enabled=True))
        print(f"  F5 打轴：第 {row} 行，帧号1={frame}")
        if self.axis_running and self._axis_setting_bool("打轴时暂停", True):
            self._pause_after_mark()
        return int(row)

    # ---- F6 手动暂停 / 恢复 ----
    def toggle_manual_pause(self) -> None:
        """F6 全局热键调用：只登记请求，暂停/恢复都交给主线程安全点。

        2026-11-09：以前热键线程会额外调一次 pause()，主线程 _manual_hold()
        又调一次；同一个切换式按钮被点两次会互相抵消（停住又被放开），
        而且两处都无条件打印「已暂停」。现在暂停只有一个控制点：
        热键线程只置 manual_pause 事件，由 _manual_hold() 点一次并校验。
        """
        if self.manual_pause.is_set():
            # 只清除请求；是否恢复由主线程的安全点决定。这里绝不能
            # 直接 resume：动作暂停中的游戏被热键线程恢复会白跑几帧。
            print("  F6：恢复请求")
            self.manual_pause.clear()
            return
        if getattr(self, "manual_resume_wait", False):
            # “暂停”操作已经停在目标帧：第一次 F6 只登记暂停请求，
            # 由主线程前置 Excel 并保持暂停；再按一次 F6 才恢复。
            print("  F6：暂停操作等待中，前置 Excel（再按 F6 恢复）")
            self.manual_pause.set()
            return
        if self.deploy_active:
            print("  F6：部署中，先回拖/取消子弹时间，再暂停")
            self.manual_pause.set()
            return
        print("  F6：暂停请求已登记，等主线程在安全点停住游戏")
        self.manual_pause.set()

    def _focus_game(self) -> bool:
        """从后台（Excel）把游戏窗口拉回前台；失败返回 False。"""
        if foreground_is(self.hwnd):
            return True
        try:
            win32gui.ShowWindow(self.hwnd, 6)   # SW_MINIMIZE
            time.sleep(0.15)
            win32gui.ShowWindow(self.hwnd, 9)   # SW_RESTORE
            time.sleep(0.25)
            win32gui.SetForegroundWindow(self.hwnd)
        except Exception:
            pass
        for _ in range(20):
            if foreground_is(self.hwnd):
                return True
            time.sleep(0.05)
        return False

    def _park_cursor(self) -> None:
        """恢复时把鼠标放到客户区左下角，避免悬停在右上角暂停/倍速按钮上。"""
        try:
            win32api.SetCursorPos(
                (self.client_left + 20, self.client_top + self.ch - 20))
        except Exception:
            pass

    def _manual_hold(self) -> float:
        """执行一次手动暂停/恢复，返回暂停时长（秒）。

        调用点必须是安全点：不会在拖干员/点技能中途把手指按住。
        """
        if not self.manual_pause.is_set():
            return 0.0
        self._raise_if_stop()
        # 唯一的暂停控制点：热键线程只登记事件，这里点一次并校验结果。
        # 暂停靠触摸注入（touch.tap 暂停按钮），必须游戏窗口在前台，否则注入
        # 被 UIPI 静默丢弃、pause() 点不中（2026-11-09 15:05 轮 F5 打轴就是
        # 前台被 Excel/控制台占着，pause 失败、游戏仍在运行）。先拉前台再点。
        if self.is_running_icon():
            if not foreground_is(self.hwnd):
                self._focus_game()
            self.pause(assume_running=True)
        if self.is_running_icon():
            self.phase = "手动暂停失败"
            print("  WARNING: F6 暂停未确认：游戏仍在运行，或右上角按钮读不到")
            print("  注入点击可能落到了 Excel/控制台上；请点一下游戏窗口，"
                  "再按 F6 恢复（本次不假装已暂停）")
        else:
            self.phase = "手动暂停中"
            print("  手动暂停：游戏已停、帧号冻结；可切后台看 Excel，再按 F6 恢复")
        # 从游戏停住就开始计时；打开 Excel 的几秒也要算进手动暂停时长，
        # 否则 pause_at 的超时会把这段时间当成卡住。
        t0 = time.perf_counter()
        rep = self.reporter
        if rep is not None and getattr(rep, "enabled", False) and \
                not getattr(rep, "started", False):
            # 开局不再为了 Excel 暂停；第一次 F6 时游戏已经停了，再打开面板。
            print("  游戏已暂停，正在打开 Excel 面板...")
            if not rep.start(wait_ready=True):
                print("  WARNING: Excel 面板没打开（文件可能还在 Excel 里开着）。"
                      "关掉占用文件后会自动重试；按 F6 可跳过面板继续")
                while self.manual_pause.is_set() and getattr(rep, "enabled", False):
                    self._raise_if_stop()
                    time.sleep(0.5)
                    if rep.start(wait_ready=True):
                        break
        if (rep is not None and getattr(rep, "enabled", False)
                and getattr(rep, "started", False)):
            print("  正在把 Excel 切到前台…")
            if rep.bring_to_front():
                print("  Excel 已前置；改完按 F6 恢复")
            else:
                print("  WARNING: Excel 没能前置，请手动切到 Excel")
        while self.manual_pause.is_set():
            self._raise_if_stop()
            time.sleep(0.05)
        while True:
            self._raise_if_stop()
            if self._focus_game():
                break
            print("  WARNING: 游戏窗口没到前台；请点一下游戏窗口，再按 F6")
            self.manual_pause.set()
            while self.manual_pause.is_set():
                self._raise_if_stop()
                time.sleep(0.05)
        self._raise_if_stop()
        self._park_cursor()
        # F6 恢复前保存并重读 Excel 里改过的未执行动作；重载期间游戏仍停着。
        # 只在轴运行中（axis_running）才做这个“原地更新”重载：它的“禁止增删
        # 行”检查只适用于执行中途改行。轴跑完后按 F6 是另一套语义——用户会
        # F5 打轴新增行，再由 _f6_service_loop 调 _resume_and_rerun 重读整个轴
        # 接着跑；那时增删行是合法的，不能再让 reload_pending_actions 的
        # “行数变了”检查把 F6 拦死（2026-11-09 07:47 轮：轴跑完 F5 加第 15 行
        # → F6 三次都 FatalReloadError，游戏卡在暂停，只能 stop）。
        if self.axis_running:
            self.reload_pending_actions()
        self._raise_if_stop()
        held = time.perf_counter() - t0
        self.phase = "运行中"
        print(f"  手动恢复（暂停了 {held:.1f}s）")
        return held

    def wait_manual_resume(self) -> None:
        """“暂停”操作：游戏已停，等用户手动恢复后再继续。

        两种恢复方式都支持：
          1. 直接点游戏里的暂停/继续按钮（is_running_icon 变回运行）；
          2. 按 F6（工具的手动暂停键），这里把它当作恢复。
        工具不会自己恢复暂停。
        """
        print("  暂停操作：等待手动恢复"
              "（点游戏暂停键，或按 F6 恢复）")
        self.manual_resume_wait = True
        try:
            while True:
                self._raise_if_stop()
                if self.is_running_icon():
                    print("  检测到游戏已手动恢复，继续后面的操作")
                    break
                if self.manual_pause.is_set():
                    # 第一次 F6：保持暂停、打开/前置 Excel，可改未执行的行；
                    # _manual_hold 会等第二次 F6 清除请求后才返回。
                    self._manual_hold()
                    self._raise_if_stop()
                    print("  收到 F6，恢复游戏")
                    self.resume()
                    continue
                time.sleep(0.05)
        finally:
            self.manual_resume_wait = False
        self.paused = False
        self._park_cursor()

    @staticmethod
    def _copy_action_fields(old: AxisAction, new: AxisAction) -> None:
        old.frame = new.frame
        old.frame2 = new.frame2
        old.kind = new.kind
        old.enabled = new.enabled
        old.oper = new.oper
        old.pos = new.pos
        old.direction = new.direction
        old.alias = new.alias
        old.note = new.note
        old.tile_pos = new.tile_pos
        old.view_side = new.view_side
        old.view_front = new.view_front
        old.rate_auto = new.rate_auto
        old.rate_percent = new.rate_percent

    def _fatal_reload(self, msg: str) -> None:
        """F6 重载失败：写 B3、停在“等待修复”，然后抛出。"""
        self.phase = "等待修复"
        if self.reporter is not None:
            self.reporter.set_error(msg)
        raise FatalReloadError(msg)

    def sort_workbook(self) -> bool:
        """把 Excel 作战记录按帧号物理排序并保存；成功返回 True。

        `load` / `run` 之前各调一次，把顺序固定下来，执行期间就不用再动表格。
        排序失败只提示、不中断（副本模式、面板没开、公式有问题都直接跳过）。
        排序会改变行号，所以成功后必须把 `row_map` 同步到动作对象，
        否则后面按 `excel_row` 写结果会落到错行。
        """
        rep = self.reporter
        if rep is None or not getattr(rep, "enabled", False) \
                or not getattr(rep, "started", False):
            return False
        path = getattr(rep, "report_path", "")
        if not path:
            return False
        if not (getattr(rep, "in_place", False)
                and hasattr(rep, "sort_records")
                and hasattr(rep, "apply_formula_rewrites")):
            return False            # 副本模式/测试替身：不动原文件
        if not rep.save_workbook():
            print("  [sort] 保存失败，跳过排序")
            return False
        try:
            pre_axis = read_axis(path)
        except Exception as exc:
            print(f"  [sort] 读 Excel 失败，跳过排序：{exc}")
            return False
        last_row = max(
            [int(a.excel_row) for a in pre_axis.actions if a.excel_row] or [2])
        try:
            order_rows, row_map = compute_row_map(pre_axis.actions, last_row)
            # 已经有序就什么都不做：run 之前会再调一次，这次调用要几乎零成本。
            if order_rows == list(range(2, 2 + len(order_rows))):
                print("  [sort] 顺序已经是目标顺序，无需移动")
                return True
            rewrites = collect_formula_rewrites(path, row_map, last_row)
        except Exception as exc:
            print(f"  [sort] 公式预检失败，跳过排序：{exc}")
            return False
        if not rep.sort_records(order_rows, last_row):
            print("  [sort] Excel 物理排序失败")
            return False
        if not rep.apply_formula_rewrites(rewrites):
            print("  [sort] 公式写回失败；表格可能已部分排序，请检查后重新 load")
            return False
        if not rep.save_workbook():
            print("  [sort] 排序后保存失败")
            return False
        # 排序后行号变了：同步到动作对象（内容没变，只是整行搬家）。
        for a in self.axis.actions:
            old_row = int(a.excel_row or 0)
            if old_row and old_row in row_map:
                a.excel_row = row_map[old_row]
        print(f"  [sort] 已按帧号排序 {len(order_rows)} 行并保存")
        return True

    def reload_pending_actions(self) -> bool:
        """F6 恢复前：保存 Excel，重读未执行的操作行，更新到现有动作对象。

        只支持改尚未执行行的 B~I；不支持中途换关卡/回费速度、插入或
        删除行、修改已执行动作。读不出的行保留旧值并给出警告。
        """
        rep = self.reporter
        if rep is None or not getattr(rep, "enabled", False) \
                or not getattr(rep, "started", False):
            return False
        report_path = getattr(rep, "report_path", "")
        if not report_path:
            return False
        print("  [reload] 保存 Excel、排序并重读…")
        if not rep.save_workbook():
            self._fatal_reload(
                "F6 重载失败：Excel 保存失败；"
                "请关闭占用文件的程序后重试（游戏保持暂停）")
        try:
            pre_axis = read_axis(report_path)
        except Exception as exc:
            self._fatal_reload(
                f"F6 重载失败：重读 Excel 出错（{exc}）；"
                "请检查文件是否被占用或损坏（游戏保持暂停）")
        if str(pre_axis.map_code or "") != str(self.axis.map_code or ""):
            self._fatal_reload(
                f"F6 重载失败：关卡代号被改（"
                f"{self.axis.map_code or '空'} → {pre_axis.map_code or '空'}）；"
                "中途不能换图，请改回后重试（游戏保持暂停）")
        # 允许在 F6 暂停期间改“打轴时暂停”，下一次 F5 立即生效。
        if "打轴时暂停" in pre_axis.settings:
            self.axis.settings["打轴时暂停"] = pre_axis.settings["打轴时暂停"]

        in_place = bool(getattr(rep, "in_place", False))
        can_sort = (hasattr(rep, "sort_records")
                    and hasattr(rep, "apply_formula_rewrites"))
        if not in_place or not can_sort:
            # 副本模式或测试替身：不物理排序，按行号更新未执行动作。
            new_by_row = {int(a.excel_row): a
                          for a in pre_axis.actions if a.excel_row}
            changed = 0
            for idx, old in enumerate(self.axis.actions):
                if idx < self.pending_from:
                    continue
                row = (int(old.excel_row)
                       if in_place and old.excel_row else 2 + idx)
                new = new_by_row.get(row)
                if new is None:
                    continue
                self._copy_action_fields(old, new)
                if in_place:
                    old.excel_row = row
                changed += 1
            print(f"  [reload] 已更新 {changed} 个未执行动作")
            return changed > 0

        last_row = max(
            [int(a.excel_row) for a in pre_axis.actions if a.excel_row] or [2])
        try:
            order_rows, row_map = compute_row_map(pre_axis.actions, last_row)
            rewrites = collect_formula_rewrites(
                report_path, row_map, last_row)
        except FormulaRewriteError as exc:
            self._fatal_reload(
                f"F6 重载失败：{exc}；请检查 J 列公式（游戏保持暂停）")
        except Exception as exc:
            self._fatal_reload(
                f"F6 重载失败：公式预检出错（{exc}）；"
                "请检查 J 列公式（游戏保持暂停）")

        if not rep.sort_records(order_rows, last_row):
            self._fatal_reload(
                "F6 重载失败：Excel 物理排序失败；"
                "请确认文件未被占用、工作表未受保护（游戏保持暂停）")
        if not rep.apply_formula_rewrites(rewrites):
            self._fatal_reload(
                "F6 重载失败：公式写回失败，表格可能已部分排序；"
                "请先检查并保存 Excel，再重新运行（游戏保持暂停）")
        if not rep.save_workbook():
            self._fatal_reload(
                "F6 重载失败：排序后保存失败；"
                "请确认文件未被占用或只读（游戏保持暂停）")

        try:
            new_axis = read_axis(report_path)
        except Exception as exc:
            self._fatal_reload(
                f"F6 重载失败：排序后重读 Excel 出错（{exc}）；"
                "请检查文件状态后重新运行（游戏保持暂停）")
        if len(new_axis.actions) != len(self.axis.actions):
            self._fatal_reload(
                f"F6 重载失败：操作行数变了（原 {len(self.axis.actions)} 行，"
                f"现 {len(new_axis.actions)} 行）；中途不能增删行，"
                "请恢复后重试（游戏保持暂停）")
        new_by_row = {int(a.excel_row): a
                      for a in new_axis.actions if a.excel_row}
        for old in self.axis.actions:
            old_row = int(old.excel_row or 0)
            new_row = row_map.get(old_row)
            new = new_by_row.get(new_row) if new_row is not None else None
            if new is None:
                self._fatal_reload(
                    f"F6 重载失败：排序后找不到原第 {old_row} 行；"
                    "请检查是否有合并单元格、空行或手工插入（游戏保持暂停）")
            self._copy_action_fields(old, new)
            old.excel_row = new_row
        print(f"  [reload] 已排序并重读 {len(self.axis.actions)} 个动作")
        return True

    def step_one_frame(self, enter_bullet: bool = False,
                       require_bullet: bool = False,
                       period: float = BULLET_STEP_PERIOD,
                       allow_blind: bool = False,
                       allow_overshoot: bool = False) -> int:
        """单帧步进：左半继续 → 等一个子弹帧 → 右半暂停。

        子弹时间固定 0.2×，1X/2X 都是 5 × 69905 / 2^21 秒。
        allow_blind=True 时，费用条多帧同宽、帧号没涨，也按“输入已
        走了一帧”继续推进模型（低回费盲步进）。
        allow_overshoot=True 时，一步走了 2 帧不当场报错，交给调用方
        判断是否还停在目标帧之前（追帧时 2 帧仍在目标前，没必要放弃）。
        """
        return self.step_one_frame_afa(
            period, require_bullet=require_bullet, allow_blind=allow_blind,
            allow_overshoot=allow_overshoot)

    def step_one_frame_afa(self, period: float | None = None,
                           require_bullet: bool = False,
                           allow_blind: bool = False,
                           allow_overshoot: bool = False) -> int:
        """AFA 同款点屏幕：左半继续 → 等 period → 右半暂停。不按 ESC/Space。"""
        before = self.current_frame()
        if period is None:
            period = BULLET_STEP_PERIOD
        # 用户要求：每次过帧最多试 5 次，5 次后帧号仍未 +1 就报错。
        for _ in range(STEP_ATTEMPTS):
            if self.manual_pause.is_set():
                self._manual_hold()
            if require_bullet and not getattr(self.counter.last, "bullet", False):
                self._dbg("  bullet 丢失，停止本次步进")
                return self.current_frame() - before
            self.touch.tap(*self.screen(*PAUSE_LEFT_RATIO))
            # 两条腿走路：
            # 1) 计数器看到 +1 → 立刻暂停。有真实相位参考，一步一个准。
            # 2) 费用条同宽、几帧都看不见变化时不能死等整拍：暂停生效还要
            #    占掉一部分帧时间，等满 166.7ms 会把下一帧一起吃进去
            #    （过 2 帧）。所以最晚 period-PAUSE_LATENCY_MARGIN 就点。
            #    提前只会 0 帧重试，绝不会过帧。
            fallback = max(0.04, period - PAUSE_LATENCY_MARGIN)
            wait_until = time.perf_counter() + fallback
            while time.perf_counter() < wait_until:
                if self.current_frame() != before:
                    break
                time.sleep(0.004)
            # 右半也必须是瞬时 tap：按住 30ms 再抬，等于把暂停生效又推后
            # 30ms（46ms 的延迟里有一大半是这 30ms 自己造的）。
            # AFA 的 ActionPauseSkill 右半同样是 Tap，没有按住。
            self.touch.tap(*self.screen(*PAUSE_RIGHT_RATIO))
            # 暂停确认等待保持完整一拍；减半会导致过冲 2 帧。
            self._sleep_precise(BULLET_PAUSE_SETTLE)
            # 右半点击可能要到下一帧才被游戏处理；等暂停图标确认后再读帧，
            # 保证 Q/E 一定是在游戏停住后按下（这里只等待，不补点击）。
            if not self.is_paused():
                deadline = time.perf_counter() + 0.1
                while time.perf_counter() < deadline:
                    time.sleep(0.01)
                    if self.is_paused():
                        break
            after = self._settled_frame()
            delta = after - before
            paused = self.is_paused()
            icon = getattr(self, "_last_icon_score", 0.0)
            state = f"paused={paused} icon={icon:.3f}"
            if delta == 1:
                self._dbg(
                    f"  screen step {before}->{after} "
                    f"period={period * 1000:.1f}ms "
                    f"settle={BULLET_PAUSE_SETTLE * 1000:.1f}ms {state}"
                )
                return 1
            if delta > 1:
                print(
                    f"  WARNING: screen step overshot {delta} frames "
                    f"{before}->{after} {state}"
                )
                if allow_overshoot:
                    # 追帧时一步走 2 帧很常见；只要还没越过目标帧，
                    # 停在 after 继续往下走就行，不要在这里判死。
                    return delta
                raise RuntimeError(
                    f"过帧过冲：{before}->{after}（{delta} 帧）")
            if (delta == 0 and allow_blind and paused
                    and (not require_bullet or self.bullet_state() is True)
                    and hasattr(self.counter, "advance_blind")):
                self.counter.advance_blind(1)
                blind_after = self.current_frame()
                self._dbg(
                    f"  blind step {before}->{blind_after} "
                    f"uncertainty={self.current_frame_hi() - blind_after} "
                    f"{state}"
                )
                return 1
            self._dbg(f"  screen step retry (0 frames) {before}->{after} {state}")
        raise RuntimeError(
            f"过帧失败：{before} 连续 {STEP_ATTEMPTS} 次没有进入下一帧")

    def step_one_frame_button(self, enter_bullet: bool = True) -> int:
        """子弹时间内用左半继续 + 右半暂停步进，不用 ESC，避免暂停菜单。"""
        self.force_pause()
        if enter_bullet and not getattr(self.counter.last, "bullet", False):
            self.touch.tap(*self.screen(*BULLET_OPER_RATIO))
            time.sleep(0.2)
        before = self.current_frame()
        for _ in range(5):
            self.touch.tap(*self.screen(*PAUSE_LEFT_RATIO))
            time.sleep(0.1)
            self.touch.tap(*self.screen(*PAUSE_RIGHT_RATIO))
            time.sleep(0.05)
            after = self.current_frame()
            delta = after - before
            if delta == 1:
                self._dbg(f"  button step {before}->{after}")
                return 1
            if delta > 1:
                print(f"  WARNING: button step overshot {delta} frames")
                return delta
            self._dbg(f"  button step retry (0 frames) {before}->{after}")
        print(f"  WARNING: button step failed at {before}->{self.current_frame()}")
        return self.current_frame() - before

    def step_one_frame_165(self) -> int:
        """子弹时间前进 166ms（AFA 顺序：ESC down → 166ms → Space down → 松开）。"""
        before = self.current_frame()
        for _ in range(10):
            key_down(VK_ESCAPE)
            self._sleep_precise(0.166)
            key_down(VK_SPACE)
            self._sleep_precise(0.05)
            key_up(VK_ESCAPE)
            key_up(VK_SPACE)
            self._sleep_precise(0.05)
            after = self.current_frame()
            delta = after - before
            if delta == 1:
                self._dbg(f"  bullet165 step {before}->{after}")
                return 1
            if delta > 1:
                print(f"  WARNING: bullet165 step overshot {delta} frames")
                return delta
            self._dbg(f"  bullet165 step retry (0 frames) {before}->{after}")
        print(f"  WARNING: bullet165 step failed at {before}->{self.current_frame()}")
        return self.current_frame() - before

    # ---- 等待 ----
    def pause_at(self, target: int, frame_threshold: int | None = None) -> int:
        """跑到目标附近并暂停，记录 trigger/landing/latency。"""
        self._raise_if_stop()
        if frame_threshold is None:
            frame_threshold = self._default_lead()
        # 上一动已经暂停在目标附近时，绝不能再 resume：240 部署后接 240 秒撤，
        # 旧代码就是在这里恢复/重暂停，白跑 6 帧到 246。
        cur = self.current_frame()
        if self.is_paused() and cur + frame_threshold >= target:
            self._dbg(f"  already paused current={cur} target={target}; keep paused")
            return cur
        # 开局“开始前暂停”开启时，manual_pause 在 run 启动前就置位了。
        # 这里必须**先**消费它，再决定要不要 resume：旧代码先 resume 把
        # 游戏点开、下一拍才发现 manual_pause 又去 _manual_hold 暂停一次，
        # 造成“恢复暂停之后又再次暂停”的抖动。开局暂停后应当一直停着，
        # 等用户 F6 恢复后才 resume 跑向目标帧。
        if self.manual_pause.is_set():
            held = self._manual_hold()
            self._raise_if_stop()
        self.resume()
        deadline = time.time() + max(10.0, max(0, target - self.current_frame()) / 30.0 * 2.0 + 5.0)
        while time.time() < deadline:
            self._raise_if_stop()
            if self.manual_pause.is_set():
                held = self._manual_hold()
                self._raise_if_stop()
                deadline += held          # 手动暂停的时长不算进超时
                self.resume()             # 游戏还是停的，继续跑向目标帧
                continue
            cur = self.current_frame()
            if cur + frame_threshold >= target:
                break
            time.sleep(0.01)
        self._raise_if_stop()
        trigger = self.current_frame()
        t_click = time.perf_counter()
        self.pause(assume_running=True)
        # 墙钟延迟：点下去到图标确认停住。lead 用它来算，跟倍速无关。
        latency_s = time.perf_counter() - t_click
        landing = self._settled_frame()
        hi = self.counter.frame_hi
        latency = landing - trigger
        self.pause_samples.append(
            (trigger, landing, latency, frame_threshold, latency_s))
        self.last_pause_trigger = trigger
        self.last_pause_landing = landing
        last = self.counter.last
        lead_ms = frame_threshold / (30.0 * self._speed()) * 1000.0
        self._dbg(
            f"  pause trigger={trigger} landing={landing}"
            f"{f'~{hi}' if hi > landing else ''} "
            f"latency={latency}帧/{latency_s * 1000:.0f}ms "
            f"lead={frame_threshold}帧/{lead_ms:.0f}ms "
            f"counter_mode={getattr(last, 'mode', None)} "
            f"warn={getattr(last, 'warn', '')}"
        )
        return landing

    def wait_until(self, target: int, frame_threshold: int | None = None) -> int:
        start = self.current_frame()
        self._dbg(f"  wait target={target} current={start}")
        landing = self.pause_at(target, frame_threshold)
        cur = landing
        self._dbg(f"  near target, current={cur} diff={cur - target}")

        guard = 0
        while self.step_enabled and self.counter.frame_hi < target:
            if self.manual_pause.is_set():
                self._manual_hold()
                continue
            self.step_one_frame()
            guard += 1
            if guard > 12:
                raise RuntimeError(f"逐帧对齐失败，停在 {self.current_frame()}")
        cur = self.current_frame()
        hi = self.counter.frame_hi
        if cur <= target <= hi and hi > cur:
            print(f"  WARNING: target {target} 落在不确定区间 {cur}~{hi}"
                  f"（费用条多帧同宽，最多差 {hi - cur} 帧；不再步进避免过冲）")
        elif cur < target:
            print(f"  WARNING: step disabled, early {target - cur} frames")
        elif cur > target:
            print(f"  WARNING: overshoot {cur - target} frames")
        return cur

    # ---- 部署 ----
    def _sleep_or_f6(self, seconds: float) -> bool:
        """等待一小段；期间按 F6 返回 True（不涉及手指按下）。"""
        end = time.perf_counter() + max(0.0, seconds)
        while time.perf_counter() < end:
            if self.manual_pause.is_set():
                return True
            time.sleep(0.01)
        return self.manual_pause.is_set()

    def _wait_drag_or_f6(self, seconds: float, pos: tuple[int, int]) -> bool:
        """拖拽途中等待；F6 时松开手指并返回 True。"""
        end = time.perf_counter() + max(0.0, seconds)
        while time.perf_counter() < end:
            if self.manual_pause.is_set():
                try:
                    self.touch.touch_up(*pos)
                except Exception:
                    pass
                print("  F6：拖拽中断，已松开手指")
                return True
            time.sleep(0.005)
        if self.manual_pause.is_set():
            try:
                self.touch.touch_up(*pos)
            except Exception:
                pass
            print("  F6：拖拽中断，已松开手指")
            return True
        return False

    def _cancel_bullet_time(self) -> None:
        """点顶部中间的取消选中/取消子弹时间位置。"""
        self.touch.click(
            *self.screen(*CANCEL_SELECTION_RATIO), hold_ms=30)
        self._dbg("  点击取消子弹时间坐标")

    def _drag_card_to_tile(self, card, middle, drop) -> bool:
        """抓卡拖到目标格；返回 True 表示被 F6 中断并已松手。"""
        self.touch.touch_down(*card)
        last = card
        if self._wait_drag_or_f6(0.05, last):
            return True
        self.touch.touch_move(*middle)
        last = middle
        if self._wait_drag_or_f6(0.3, last):
            return True
        for i in range(1, DEPLOY_DRAG_STEPS + 1):
            last = (
                int(middle[0] + (drop[0] - middle[0]) * i / DEPLOY_DRAG_STEPS),
                int(middle[1] + (drop[1] - middle[1]) * i / DEPLOY_DRAG_STEPS),
            )
            self.touch.touch_move(*last)
            if self._wait_drag_or_f6(DEPLOY_DRAG_STEP, last):
                return True
        if self._wait_drag_or_f6(0.15, last):
            return True
        if self.manual_pause.is_set():
            try:
                self.touch.touch_up(*last)
            except Exception:
                pass
            print("  F6：拖拽中断，已松开手指")
            return True
        self.touch.touch_up(*drop)
        return False

    def _drag_direction(self, tile, dx: int, dy: int) -> bool:
        """从格子中心拖方向。

        F6 时**不立刻松手**：按原速沿已走过路径拖回起点，在起点松手，
        由调用方点击取消子弹时间坐标。返回 True 表示已回拖并松手。
        """
        dir_pos = (tile[0] + dx, tile[1] + dy)
        points = [tile]
        self.touch.touch_down(*tile)

        def wait_f6(seconds: float) -> bool:
            end = time.perf_counter() + max(0.0, seconds)
            while time.perf_counter() < end:
                if self.manual_pause.is_set():
                    return True
                time.sleep(0.005)
            return self.manual_pause.is_set()

        def return_to_start() -> None:
            # 不松手，沿 visited 反向逐点拖回；每点等待与去程相同。
            for p in reversed(points[:-1]):
                self.touch.touch_move(*p)
                time.sleep(DEPLOY_DRAG_STEP)
            self.touch.touch_up(*points[0])
            print("  F6：方向拖拽中断，按原速拖回起点并松手")

        for i in range(1, DEPLOY_DRAG_STEPS + 1):
            p = (
                int(tile[0] + dx * i / DEPLOY_DRAG_STEPS),
                int(tile[1] + dy * i / DEPLOY_DRAG_STEPS),
            )
            self.touch.touch_move(*p)
            points.append(p)
            if wait_f6(DEPLOY_DRAG_STEP):
                return_to_start()
                return True
        if wait_f6(0.15):
            return_to_start()
            return True
        if self.manual_pause.is_set():
            return_to_start()
            return True
        self.touch.touch_up(*dir_pos)
        return False

    def deploy(self, action: AxisAction,
               resume_after: bool = False) -> int | None:
        if not action.oper or action.view_side is None or not action.direction:
            raise ValueError(f"部署信息不完整: {action}")

        # 用户流程：提前暂停 → 点待部署区最右边的卡进入子弹时间
        # → AFA“前进 166ms”到部署帧 → 再选中目标干员拖到格子。
        self._click_rightmost_deploy_card()
        time.sleep(0.25)
        self._dbg(f"  enter bullet frame={self.current_frame()} "
                  f"bullet={getattr(self.counter.last, 'bullet', None)}")
        if self.bullet_state() is not True:
            print("  WARNING: 部署卡没有进入子弹时间，判定部署未成功")
            return None
        guard = 0
        while self.step_enabled and self.current_frame() < action.frame:
            if self.bullet_state() is not True:
                print("  WARNING: 部署步进前丢失子弹时间，判定部署未成功")
                return None
            self.step_one_frame(require_bullet=True)
            guard += 1
            if guard > 12:
                print("  WARNING: 部署过帧失败，判定部署未成功")
                return None
        cur = self.current_frame()
        if cur < action.frame:
            print(f"  WARNING: deploy early {action.frame - cur} frames")
        if cur > action.frame:
            # 目标帧已经过了就不要再拖：晚一帧部署可能“成功”，但这不是
            # 轴要的 30 帧结果，会掩盖 30 帧应有的失败。取消选中后按失败记。
            print(f"  WARNING: deploy overshoot {cur - action.frame} frames；"
                  f"判定部署未成功")
            try:
                self.touch.click(
                    *self.screen(*CANCEL_SELECTION_RATIO), hold_ms=30)
            except Exception:
                pass
            return None

        # 目标选择/拖拽/方向整体可重做。只有“已放到目标格之后”的
        # F6 才由主线程按顺序处理（取消/回拖/暂停）；拖卡过程中仍由
        # 热键线程先暂停，然后松手重做。
        while True:
            if self.manual_pause.is_set():
                self._manual_hold()
            # F6 暂停可能让游戏退出 bullet；重新进一次子弹再拖。
            if self.bullet_state() is not True:
                self._dbg("  bullet lost before target; re-enter by right card")
                if not self._enter_bullet_by_card():
                    print("  WARNING: 重新进子弹失败，判定部署未成功")
                    return None

            frame = self.stream.latest()
            if frame is None:
                raise RuntimeError("没有抓屏帧")
            ax, ay, score = locate_avatar(frame, action.oper)
            self._dbg(f"  avatar {action.oper} at ({ax:.3f},{ay:.3f}) "
                      f"score={score:.3f}")
            card = self.screen(ax, ay)
            self.touch.click(*card)
            self.selected_oper = action.oper
            if self._sleep_or_f6(0.3):
                self._manual_hold()
                continue
            # 选中后待部署区会滚动，卡片位置会变；必须重新找一次，
            # 否则会从旧位置抓起相邻干员（prts-plus 也在这里重新 locate）。
            try:
                frame = self.stream.latest()
                if frame is not None:
                    ax, ay, score = locate_avatar(frame, action.oper)
                    self._dbg(f"  avatar after select at ({ax:.3f},{ay:.3f}) "
                              f"score={score:.3f}")
                    card = self.screen(ax, ay)
            except RuntimeError as exc:
                print(f"  WARNING: 选中后找不到头像，沿用旧位置: {exc}")
            if self.manual_pause.is_set():
                self._manual_hold()
                continue

            tile = self.screen(*action.view_side)
            # prts-plus 的落点比格子中心低 2%，方向拖拽仍从格子中心开始。
            drop = (tile[0], tile[1] + int(0.02 * self.ch))
            middle = (card[0], card[1] - int(0.03 * self.ch))
            # 抓起卡片后先停在中间，等镜头转到部署视角，再移向目标格。
            # 镜头旋转需要约 0.3s；旋转途中落点会偏一格。
            if self._drag_card_to_tile(card, middle, drop):
                self._manual_hold()
                continue

            # 从这里开始已经放到目标格：F6 要按“取消子弹/回拖”处理，
            # 所以让热键线程只设请求、不抢先点暂停。
            self.deploy_active = True
            try:
                # 用户要求：放到目标格后先等约 0.05 秒判定部署结果，
                # 此时还没调方向。成功时仍在 bullet；30 帧那次这里是 False。
                time.sleep(0.05)
                drop_bullet = self.bullet_state()
                self._dbg(f"  drop tile bullet={drop_bullet}")
                if drop_bullet is not True:
                    print("  WARNING: 放到目标格后不在子弹时间，"
                          "判定部署未成功")
                    return None

                # 拖到目标格后、还没开始调方向时按 F6：
                # 点取消子弹坐标退出，等用户恢复；仍在指定帧就重新部署。
                if self.manual_pause.is_set():
                    print("  F6：已放到目标格但未调方向，"
                          "取消子弹时间后暂停")
                    self._cancel_bullet_time()
                    self._manual_hold()
                    if self.current_frame() == action.frame:
                        print("  F6 恢复后仍在指定帧，重新部署")
                        continue
                    print("  WARNING: F6 恢复后已不在指定帧，不重做部署")
                    return None

                direction = action.direction
                dx = dy = 0
                if direction == "右":
                    dx = int(DIRECTION_DX * self.cw)
                elif direction == "左":
                    dx = -int(DIRECTION_DX * self.cw)
                elif direction == "上":
                    dy = -int(DIRECTION_DY * self.ch)
                elif direction == "下":
                    dy = int(DIRECTION_DY * self.ch)
                if dx or dy:
                    # 方向拖拽中按 F6：helper 会不松手、按原速拖回起点
                    # 再松手；这里点取消子弹、暂停，仍在指定帧就重做部署。
                    if self._drag_direction(tile, dx, dy):
                        print("  F6：方向拖拽中断，取消子弹时间后暂停")
                        self._cancel_bullet_time()
                        self._manual_hold()
                        if self.current_frame() == action.frame:
                            print("  F6 恢复后仍在指定帧，重新部署")
                            continue
                        print("  WARNING: F6 恢复后已不在指定帧，"
                              "不重做部署")
                        return None
                    time.sleep(0.15)
            finally:
                self.deploy_active = False
            break

        time.sleep(0.2)
        deploy_frame = self.current_frame()
        print(f"  deployed, frame={deploy_frame}")
        if resume_after:
            # 下一个操作离得远：指针松开后直接恢复，让战斗正常跑。
            self.resume()
        return deploy_frame

    # ---- 技能 / 撤退 ----
    def select_while_paused(self, rx: float, ry: float,
                            pause_mode: str = "hold") -> None:
        """AFA 式三连点：继续 -> 点干员 -> 暂停。

        对齐 AFA 的 ActionPauseSkill/ActionPauseRetreat：三下都是
        “按下+抬起”的瞬时 Tap，背靠背不加间隔，右半不按住；
        点完只等 ClickDelay（50ms），再交给调用方按 Q/E。
        """
        if self.manual_pause.is_set():
            self._manual_hold()
        left = self.screen(*PAUSE_LEFT_RATIO)
        target = self.screen(rx, ry)
        right = self.screen(*PAUSE_RIGHT_RATIO)
        self._dbg(f"  select paused({pause_mode}): left={left} target={target} right={right}")
        if pause_mode == "direct":
            # 暂停中直接点干员：选不中，只当前置唤醒点击用。
            self.touch.tap(*target)
            time.sleep(W_SETTLE)
            return
        # 完全照 AFA：三次 tap 背靠背，中间不加间隔。
        # 记录每一次点击的时刻和当时帧号：三连点只要有一次跨帧，
        # 日志就能看出是哪一段（继续→选中 还是 选中→暂停）吃掉了帧。
        t0 = time.perf_counter()
        f0 = self.current_frame()
        self.touch.tap(*left)
        t1 = time.perf_counter()
        f1 = self.current_frame()
        self.touch.tap(*target)
        t2 = time.perf_counter()
        f2 = self.current_frame()
        t3 = f3 = None
        if pause_mode == "tap":
            self.touch.tap(*right)
            # 选中/子弹时间流程不随游戏倍速变快，等待固定 1.5 拍；
            # 用 _frame_period() 会在 2X 下被缩短一半，导致选人失败。
            time.sleep(1.5 / 30.0)
            t3, f3 = time.perf_counter(), self.current_frame()
        elif pause_mode == "hold":
            # 照 AFA：第三下也是瞬时 Tap，不按住。按住 30ms 会让“暂停”
            # 晚一拍才被游戏认下，正好把选人挤到下一帧。
            self.touch.tap(*right)
            t3, f3 = time.perf_counter(), self.current_frame()
        elif pause_mode == "center":
            self._click_pause()  # 暂停按钮中心，50ms
            time.sleep(0.05)
        elif pause_mode == "key_first":
            # 先不暂停：让调用方立刻按 Q/E，再马上补暂停。
            time.sleep(0.03)
        elif pause_mode == "settle":
            # 等游戏处理完“点中干员”，再点暂停；否则暂停会被忽略到动画结束。
            time.sleep(0.05)
            self._click_pause()
            time.sleep(0.05)
        elif pause_mode == "esc":
            key_down(VK_ESCAPE)
            time.sleep(0.05)
            self._esc_held = True
        else:
            raise ValueError(f"未知 pause_mode: {pause_mode}")
        # 照 AFA：点完只等 ClickDelay（默认 50ms）就按键。
        # 先把这 50ms 等完再打印，避免写日志的 I/O 挤在点击和按键之间。
        time.sleep(CLICK_DELAY)
        if t3 is not None:
            self._dbg(
                f"  tap3 {pause_mode}: start={f0} "
                f"left@{(t1 - t0) * 1000:.1f}ms {f0}->{f1} "
                f"target@{(t2 - t1) * 1000:.1f}ms {f1}->{f2} "
                f"pause@{(t3 - t2) * 1000:.1f}ms {f2}->{f3}")

    def _send_action_key(self, vk: int, label: str) -> None:
        """普通短按 Q/E；按完由调用方等待 2 秒再继续。"""
        if self.manual_pause.is_set():
            self._manual_hold()
        before = self.current_frame()
        key_down(vk)
        time.sleep(0.05)
        key_up(vk)
        self._dbg(f"  {label} key frame={before}->{self.current_frame()}")

    def skill_or_retreat(self, action: AxisAction) -> None:
        if action.view_front is None:
            raise ValueError(f"技能/撤退没有位置: {action}")
        self.force_pause()
        self.select_while_paused(*action.view_front, pause_mode=self.select_pause_mode)

        vk = RETREAT_KEY if action.kind == "撤退" else SKILL_KEY
        self._send_action_key(vk, action.kind)
        if self._esc_held:
            time.sleep(0.05)
            key_up(VK_ESCAPE)
            self._esc_held = False
        time.sleep(0.05)
        # Q 后游戏本来就暂停；不要额外按 ESC，只等镜头复位。
        time.sleep(CAMERA_SETTLE)
        print(f"  {action.kind} done, frame={self.current_frame()}")

    def select_and_key_fast(self, action: AxisAction,
                            assume_paused: bool = True, index: int = -1) -> None:
        """连撤用：选中后立刻按 Q/E，不在这里暂停。"""
        if action.view_front is None:
            raise ValueError(f"技能/撤退没有位置: {action}")
        if assume_paused:
            self.touch.tap(*self.screen(*PAUSE_LEFT_RATIO))
        self.touch.tap(*self.screen(*action.view_front))
        time.sleep(0.03)
        vk = RETREAT_KEY if action.kind == "撤退" else SKILL_KEY
        self._send_action_key(vk, f"[{index}] {action.kind} {action.oper or ''}")

    def batch_skill_retreat(self, group: list[AxisAction], start_index: int = 0,
                            resume_after: bool = False) -> None:
        """连技/连撤：选中并确认 bullet → 步进 → 按 E/Q 并确认生效。

        真实轴会遇到技能没放出（E 没反应）或撤退没成功（Q 没反应），
        按完仍停在子弹时间。这里每次选中和按键都要确认；连续失败就
        报错停下，绝不带着错位继续跑后面的动作。
        """
        for k, action in enumerate(group):
            # 连技/连撤的每个动作之间是安全点：F6 在这里才响应，
            # 不会在选中/按键中途把手指按住。
            if self.manual_pause.is_set():
                self._manual_hold()
            if action.view_front is None:
                raise ValueError(f"技能/撤退没有位置: {action}")
            idx = start_index + k
            self.current_action = (idx, action)
            if self.reporter:
                self.reporter.action_start(idx, action)

            self._dbg(f"  [{idx}] enter bullet by right card at "
                      f"frame={self.current_frame()}")
            pre_w = self.current_frame()
            entered_by_card = self._enter_bullet_by_card()
            if not entered_by_card:
                print("  WARNING: 最右待部署卡没进子弹，退回点目标选择")
                self._select_operator(action)
                self._last_deploy_card_screen = None
            post_w = self.current_frame()
            self._dbg(
                f"  entered bullet frame {pre_w}->{post_w} "
                f"bullet={self.bullet_state()}"
            )

            guard = 0
            while self.step_enabled and self.current_frame() < action.frame:
                if self.bullet_state() is not True:
                    # 绝不能在没有子弹时间时用 80ms 步进：1X 下会跨 3~4 帧。
                    self._dbg("  bullet lost before step; re-enter by right card")
                    entered_by_card = self._enter_bullet_by_card()
                    if not entered_by_card:
                        self._select_operator(action)
                        self._last_deploy_card_screen = None
                self.step_one_frame(require_bullet=True)
                guard += 1
                if guard > 12:
                    raise RuntimeError(f"过帧失败，停在 {self.current_frame()}")
            cur = self.current_frame()
            if cur < action.frame:
                print(f"  WARNING: step disabled, early {action.frame - cur} frames")
            elif cur > action.frame:
                raise RuntimeError(
                    f"步进过冲：目标 {action.frame}，当前 {cur}；"
                    f"已停止，避免在错误帧按 E/Q")

            # 到目标帧后，再点一次刚才选中的部署卡，退出子弹/部署选择。
            exited_card = False
            if entered_by_card and self._last_deploy_card_screen is not None:
                self._dbg(f"  click selected deploy card again to exit bullet "
                          f"at frame={self.current_frame()}")
                self.touch.click(*self._last_deploy_card_screen, hold_ms=30)
                exited_card = True
                if self._sleep_or_f6(0.2):
                    self._manual_hold()

            # 退出部署选择后，再点场上已部署的目标，确认选中后按 Q/E。
            # 只有没走“再点部署卡退出”的兜底路径才需要重新进 bullet。
            if (not exited_card and self.bullet_state() is not True):
                self._dbg("  bullet lost before target select; re-enter")
                entered_by_card = self._enter_bullet_by_card()
                if not entered_by_card:
                    self._select_operator(action)
                    self._last_deploy_card_screen = None
            self._dbg(f"  select battlefield target at frame="
                      f"{self.current_frame()}")
            self.select_while_paused(*action.view_front, pause_mode="direct")
            self.selected_oper = action.oper
            if self.bullet_state() is not True:
                print("  WARNING: 点场上目标后失去子弹时间，退回原选中流程")
                self._select_operator(action)
            key_frame = self.current_frame()
            last = self.counter.last
            self._dbg(
                f"  key frame={key_frame} "
                f"counter_t={getattr(last, 't', 0.0):.3f} "
                f"counter_mode={getattr(last, 'mode', None)} "
                f"warn={getattr(last, 'warn', '')}"
            )
            if action.kind == "撤退":
                ok = self._send_action_verified(
                    action, RETREAT_KEY, "撤退 Q", "撤退")
            else:
                ok = self._send_action_verified(
                    action, SKILL_KEY, "技能 E", "技能")
            key_frame = self.current_frame()
            self._dbg(
                f"  after key: paused={self.is_paused()} "
                f"bullet={self.bullet_state()} ok={ok} "
                f"icon={getattr(self, '_last_icon_score', 0.0):.3f} "
                f"frame={key_frame}"
            )
            time.sleep(ACTION_SETTLE)
            if action.kind == "技能":
                # 技能没开启时镜头不会复位；等动作结束后点顶部中间取消选中。
                # 成功/失败都取消；失败不补帧，只记录。
                self.touch.click(
                    *self.screen(*CANCEL_SELECTION_RATIO), hold_ms=30)
                print("  取消选中，镜头复位")
            elif not ok:
                # 撤退失败后仍停在 bullet；取消选中，避免影响下一动作。
                self.touch.click(
                    *self.screen(*CANCEL_SELECTION_RATIO), hold_ms=30)
                print("  撤退失败，取消选中，准备下一动作")
                time.sleep(0.3)
            self._dbg(
                f"  after settle: paused={self.is_paused()} "
                f"bullet={self.bullet_state()} frame={self.current_frame()}"
            )
            if self.reporter:
                note = "OK" if ok else "失败（仍在子弹时间）"
                actual = key_frame if ok else None
                self.reporter.action_done(
                    idx, action, action.frame, actual, note)
        self.current_action = None
        if resume_after:
            # 这一组之后下一个操作离得远：E/Q 完直接恢复游戏。
            self.resume()

    # ---- 旧主循环（保留作对照，不再调用） ----
    def _run_legacy(self, max_actions: int = 1, start_action: int = 0) -> None:
        selected = self.axis.actions[start_action:start_action + max_actions]
        i = 0
        while i < len(selected):
            self.pending_from = start_action + i
            if self.manual_pause.is_set():
                self._manual_hold()
            action = selected[i]
            self.selected_oper = None
            if action.kind == "暂停":
                # 轴里的“暂停”行：提前暂停 → 点最右待部署卡进子弹 →
                # 步进到指定帧 → 再点该卡退出子弹 → 等用户手动恢复。
                # 这一行不需要干员/坐标/朝向。
                idx0 = start_action + i
                print(f"[{idx0}] frame={action.frame} 暂停，等待手动恢复")
                self.current_action = (idx0, action)
                if self.reporter:
                    self.reporter.action_start(idx0, action)

                self.pause_at(max(0, action.frame - PRE_STEP_FRAMES))
                ok = self._enter_bullet_by_card()
                if ok:
                    self.selected_oper = "待部署卡"
                guard = 0
                while ok and self.current_frame() < action.frame:
                    if self.bullet_state() is not True:
                        self._dbg("  pause-op: bullet lost, re-enter by right card")
                        ok = self._enter_bullet_by_card()
                        if not ok:
                            break
                    self.step_one_frame(require_bullet=True)
                    guard += 1
                    if guard > 12:
                        print("  WARNING: 暂停操作过帧失败")
                        ok = False
                        break
                cur = self.current_frame()
                if cur != action.frame:
                    print(f"  WARNING: 暂停操作未对准目标帧 "
                          f"target={action.frame} current={cur}")
                    ok = False

                # 退出部署卡选择/子弹时间，回到普通暂停。
                if self._last_deploy_card_screen is not None:
                    self.touch.click(*self._last_deploy_card_screen,
                                     hold_ms=30)
                    if self._sleep_or_f6(0.2):
                        self._manual_hold()
                if self.bullet_state() is True:
                    self._cancel_bullet_time()
                    time.sleep(0.2)
                if self.is_running_icon():
                    self.pause()
                landed = self.current_frame()
                self.selected_oper = None
                actual = (landed if (ok and self.is_paused()
                                     and self.bullet_state() is not True
                                     and landed == action.frame) else None)
                if actual is None and ok:
                    print(f"  WARNING: 暂停操作落点 {landed} "
                          f"!= 目标 {action.frame}")
                if self.reporter:
                    note = "OK" if actual is not None else "失败（暂停未成功）"
                    self.reporter.action_done(
                        idx0, action, action.frame, actual, note)
                self.current_action = None
                if actual is not None:
                    self.wait_manual_resume()
                else:
                    print("  WARNING: 暂停操作没有停住游戏，继续后面的动作")
                i += 1
                continue
            group = [action]
            j = i + 1
            if action.kind in ("技能", "撤退"):
                while (j < len(selected)
                       and selected[j].kind in ("技能", "撤退")
                       and selected[j].frame - group[-1].frame <= BATCH_WINDOW):
                    group.append(selected[j])
                    j += 1
            for k, item in enumerate(group):
                print(f"[{start_action + i + k}] frame={item.frame} {item}")
            # 下一个操作离得远（不在连续窗口内）时，动作完成后直接恢复游戏；
            # 挨得近就保持暂停，避免恢复/再暂停之间丢帧。
            next_frame = selected[j].frame if j < len(selected) else None
            resume_after = (next_frame is None
                            or next_frame - group[-1].frame > BATCH_WINDOW)
            if next_frame is None and self.end_pause:
                # 轴设置“结束后暂停”：最后一组也不要提前 resume。
                resume_after = False
            if action.kind in ("技能", "撤退"):
                # 和 prts-plus 一样提前几帧暂停：先选中干员进入子弹时间，
                # 再逐帧走到目标帧，最后按 Q/E。不要等到目标帧才选人。
                # 这里只暂停、不追求精确落点；差几帧交给子弹时间步进补齐，
                # 避免在 1 倍速下做不可靠的短步进。
                self.pause_at(max(0, group[0].frame - PRE_STEP_FRAMES))
                self.batch_skill_retreat(group, start_index=start_action + i,
                                         resume_after=resume_after)
            else:
                # 部署也改成“提前暂停 → 点最右卡进子弹 → AFA 步进到目标”。
                idx0 = start_action + i
                self.current_action = (idx0, action)
                if self.reporter:
                    self.reporter.action_start(idx0, action)
                self.pause_at(max(0, group[0].frame - PRE_STEP_FRAMES))
                if action.kind == "部署":
                    actual = self.deploy(action, resume_after=resume_after)
                    if self.reporter:
                        note = ("OK" if actual is not None
                                else "失败（部署未成功）")
                        self.reporter.action_done(
                            idx0, action, action.frame, actual, note)
                else:
                    raise NotImplementedError(f"未知操作: {action.kind}")
            self.current_action = None
            self.selected_oper = None
            i = j
        # 最后一个动作后用户按了 F6：先把暂停等完，再决定要不要恢复。
        if self.manual_pause.is_set():
            print("  轴已完成，但手动暂停仍生效：游戏保持暂停")
            try:
                self._manual_hold()
            except StopRequested:
                self.manual_pause.clear()
                print("  轴已停止")
                return
        if self.end_pause:
            if self.is_running_icon():
                self.pause()
            print("  轴完成：结束后暂停（轴设置），游戏保持暂停")
            if self.reporter is not None:
                if self.reporter.bring_to_front():
                    print("  轴完成：Excel 已前置")
                else:
                    print("  WARNING: 轴完成，但 Excel 没能前置")
        else:
            # 轴跑完后取消暂停，让战斗继续。
            self.resume()

    # ---- 新：按帧推进 ----
    def _enter_bullet_chain(self, action: AxisAction | None = None) -> bool:
        """进子弹：最右待部署卡 → 场上目标 → 部署栈倒序的存活干员。"""
        if self.bullet_state() is True:
            # 已经在子弹里就别再点了：再点一下会切换/取消选中，
            # 也容易因为“选中的卡没旗标”点到右数第二张。
            self._dbg("  bullet entry: 已在子弹时间，跳过")
            return True
        self.phase = "进入子弹时间"
        # 先点固定点（卡片行右对齐，最右那张一直在这儿），不做任何识别：
        # 识别会在卡被选中/动画时退到右数第二张，点下去把选中切走。
        fixed = self.screen(*DEPLOY_CARD_RATIO)
        self._dbg(f"  bullet entry: 固定卡片点 {fixed}")
        self._last_deploy_card_screen = fixed
        self.touch.click(*fixed, hold_ms=30)
        if self._wait_bullet(True, BULLET_CONFIRM):
            self._bullet_entry_card_screen = fixed
            return True
        # 固定点没反应（卡片行布局变了 / 最右卡刚好被用掉）才退回旗标找卡。
        self._dbg("  bullet entry: 固定点没进子弹，改用旗标找卡")
        self._click_rightmost_deploy_card()
        if self._wait_bullet(True, BULLET_CONFIRM):
            self._bullet_entry_card_screen = self._last_deploy_card_screen
            return True
        self._bullet_entry_card_screen = None

        targets: list[tuple[str | None, tuple[float, float]]] = []
        # 再确认一次：刚才那一下可能已经选中了卡，只是视觉样本慢了半拍。
        # 如果这时真的在子弹里，还去点场上干员就等于把手里的卡**部署出去**。
        if self._wait_bullet(True, 0.15):
            self._dbg("  bullet entry: 延迟确认到子弹时间，不再点场上单位")
            self._bullet_entry_card_screen = self._last_deploy_card_screen
            return True
        if (action is not None and action.kind in ("技能", "撤退")
                and action.view_front is not None):
            targets.append((action.oper, action.view_front))
        for rec in reversed(self.deployed_stack):
            view = rec.get("view_front")
            name = rec.get("name")
            if view is None:
                continue
            if any(n == name for n, _ in targets):
                continue
            targets.append((name, view))

        for name, view in targets:
            self._dbg(f"  bullet entry: select unit {name}")
            if self.is_running_icon():
                self.pause(assume_running=True)
            # 先做一次前置轻点再走三连点。`_select_operator` 就是这么排的；
            # 去掉它三连点的“继续”容易被客户端吞掉（实测 5/5 全失败）。
            self.select_while_paused(*view, pause_mode="direct")
            self.select_while_paused(*view, pause_mode=self.select_pause_mode)
            if self._wait_bullet(True, BULLET_CONFIRM):
                self.selected_oper = name
                return True
            if self.is_running_icon():
                self.pause()
        print("  WARNING: 所有入口都没进入子弹时间")
        return False

    def _wait_bullet_exit(self, since_t: float | None,
                          confirm: float | None = None) -> bool:
        """点击后等退出子弹。

        只看时间不够：bullet 字段是视觉线程算出来的，点击后 20ms 内很可能
        还没产出新样本，读到的仍是旧状态。所以先看时间，再看样本时间戳
        `last.t` 是否推进了；推进了还显示 bullet，才算真的没退出去。
        `confirm` 可单独放大某一步的确认窗口（如“点最右卡”这条高频路径）。
        """
        soft = time.perf_counter() + (
            EXIT_CONFIRM if confirm is None else confirm)
        hard = soft + EXIT_CONFIRM_GRACE
        while True:
            if self.bullet_state() is not True:
                return True
            now = time.perf_counter()
            if now >= hard:
                return False
            if now >= soft:
                t = getattr(self.counter.last, "t", None)
                if since_t is None or t is None:
                    return False
                if (t - since_t) >= EXIT_FRESH_MIN:
                    # 看到点击之后的新样本了，状态确实是 bullet → 真的没退出
                    return False
            time.sleep(0.005)

    def _exit_bullet_chain(self) -> bool:
        """退出子弹：最右卡 → 退出点 → 同一张最右卡。

        手动操作时三次点击都落在同一张最右卡上；旧实现每次重新找最右，
        卡片滚动/选中后会点到不同的卡，反而退不出去。这里先固定当前
        最右卡的头像中心，第一次和第三次都点它。
        """
        self.phase = "退出子弹时间"

        def exited() -> bool:
            return self.bullet_state() is not True

        if exited():
            return True
        # 部署后的卡片行可能还在收动画；等它稳定再取最右头像中心。
        time.sleep(0.15)
        current_pt, current_card = self._rightmost_card_screen()
        cached_pt = self._last_deploy_card_screen
        if cached_pt is not None:
            # 最右卡被选中时旗标会消失，重新检测会掉到“右数第二张”。
            # 优先用进子弹时真正点中的那张最右卡坐标。
            first_pt = cached_pt
            self._dbg(f"  退出子弹：优先用进子弹时的最右卡位置 {first_pt}")
            if current_pt is not None and current_pt != first_pt:
                self._dbg(f"  退出子弹：当前旗标检测到的是 {current_pt}，先不使用")
        elif current_pt is not None:
            first_pt = current_pt
            ax, ay = current_card["avatar_center"]
            self._dbg(f"  退出子弹：最右卡头像中心 "
                      f"({ax / 1280.0:.3f},{ay / 720.0:.3f})")
        else:
            first_pt = self.screen(*BULLET_OPER_RATIO)
            self._dbg(f"  退出子弹：找不到卡，改用固定位置 {first_pt}")

        def attempt(fixed: bool) -> bool:
            card_pt = first_pt
            # 先点哪个，取决于“现在选中的是什么”——这个状态我们知道：
            #   _bullet_entry_card_screen 有值 = 是用点最右卡进的子弹，
            #   现在选中的是待部署卡 → 点同一张卡直接退出。
            #   没有值 = 选中的是场上干员 → 点退出点直接退出。
            # 2026-11-08 实测：卡选中态先点退出点会白花 80ms，
            # 一次退子弹 272ms → 437ms。所以按状态选，两边都是一下退出。
            if self._bullet_entry_card_screen is not None:
                order = ("card", "exit", "card")
            else:
                order = ("exit", "card", "card")
            for step in order:
                if exited():
                    return True
                since_t = getattr(self.counter.last, "t", None)
                if step == "exit":
                    self._click_exit_bullet_point()
                else:
                    if not fixed:
                        # 手动操作时可能每次点“当前最右卡”；卡片行
                        # 如果滚动，重新取头像中心比死守旧坐标更稳。
                        now_pt, now_card = self._rightmost_card_screen()
                        if now_pt is not None:
                            card_pt = now_pt
                            ax, ay = now_card["avatar_center"]
                            self._dbg(f"  退出子弹：当前最右卡头像中心 "
                                      f"({ax / 1280.0:.3f},{ay / 720.0:.3f})")
                    else:
                        self._dbg("  退出子弹：点同一张最右卡")
                    self.touch.click(*card_pt, hold_ms=30)
                # 点卡这一步成功概率最高，而且“点中了但没看见”会让我们白点
                # 下一次（实测退子弹 337ms ≈ 2 帧）。所以点卡给 120ms 确认；
                # 点退出点只作快速试探，给 80ms 就够。
                if self._wait_bullet_exit(since_t,
                                          0.08 if step == "exit" else 0.12):
                    self._dbg("  已退出子弹时间")
                    return True
            return exited()

        if attempt(fixed=True):
            return True
        self._dbg("  固定同一张卡没退出，改为每次重新找最右卡")
        if attempt(fixed=False):
            return True
        print("  WARNING: 六次都没退出子弹时间")
        return exited()

    def _log_pause(self, action, target: int, lead: int,
                   trigger: int | None, landing: int | None) -> None:
        entry = {
            "time": time.strftime("%Y-%m-%d %H:%M:%S"),
            "index": getattr(action, "_index", -1),
            "kind": getattr(action, "kind", "") or "",
            "target": target,
            "trigger": trigger,
            "landing": landing,
            "latency": (landing - trigger
                        if trigger is not None and landing is not None
                        else None),
            "lead": lead,
            # lead 的原始单位是秒；帧数只是按当前倍速换算出来的。
            "lead_s": round(lead / (30.0 * self._speed()), 3),
        }
        self.pause_log.append(entry)
        print(f"  pause record: [{entry['index']}] {entry['kind']} "
              f"target={target} trigger={trigger} landing={landing} "
              f"lead={lead}帧/{entry['lead_s'] * 1000:.0f}ms")

    def _write_pause_log(self) -> None:
        rows = getattr(self, "pause_log", [])
        if not rows:
            return
        os.makedirs("out", exist_ok=True)
        path = os.path.join("out", "pause_log.csv")
        fields = ["time", "index", "kind", "target",
                  "trigger", "landing", "latency", "lead", "lead_s"]
        with open(path, "w", newline="", encoding="utf-8-sig") as fh:
            writer = csv.DictWriter(fh, fieldnames=fields)
            writer.writeheader()
            writer.writerows(rows)
        print(f"  暂停记录已写入 {path}")

    def _settle_paused(self) -> bool:
        if self.bullet_state() is True:
            if not self._exit_bullet_chain():
                return False
        if self.is_running_icon():
            self.pause(assume_running=True)
        return True

    def _prepare_to_frame(self, action: AxisAction | None, target: int) -> bool:
        """跑到 target-lead 暂停 → 进子弹 → 步进到 target；停在子弹时间。"""
        self._raise_if_stop()
        self.phase = f"等待目标帧 {target}"
        self._bullet_entry_card_screen = None
        cur = self.current_frame()
        # 计数器回零/重开保护：单次 run 里帧号正常只增不减，倒退上百帧
        # 只可能是离开战斗或重开。以前会从 0 一路步进到上千帧，白跑几分钟。
        seen = max(getattr(self, "_max_frame_seen", 0), cur)
        self._max_frame_seen = seen
        if cur + FRAME_ROLLBACK_LIMIT < seen:
            print(f"  WARNING: 帧号从 {seen} 回退到 {cur}，可能已离开战斗/"
                  f"重开；停止本次轴，请重新 run")
            self.stop_requested.set()
            return False
        if cur > target:
            print(f"  WARNING: 已过目标帧 target={target} current={cur}")
            return False
        # 严格不晚帧：提前量按实测延迟（秒）反推，再换算成帧。
        lead_s = self._strict_pause_lead_s()
        # 轴里填 0ms 就允许 0 帧提前量，不再用 max(1, ...) 兜底。
        lead = max(0, int(round(lead_s * 30.0 * self._speed())))
        trigger = landing = cur
        if target - cur > lead:
            # pause_at 的条件是 “cur + threshold >= target”，它内部的 target
            # 就是目标帧；以前传 target-lead 等于把 lead 减了两次，触发点
            # 变成 target-2*lead，白走 4~5 步（2026-11-08 实测确认）。
            self.pause_at(target, frame_threshold=lead)
            trigger = self.last_pause_trigger
            landing = self.last_pause_landing
        elif self.is_running_icon():
            # 距离不够提前量：立即暂停，再用子弹时间补帧。
            trigger = cur
            self.pause(assume_running=True)
            landing = self._settled_frame()
        self._log_pause(action, target, lead, trigger, landing)
        if landing is not None and landing > target:
            print(f"  WARNING: 暂停落点已晚于目标 target={target} "
                  f"landing={landing}；严格模式下放弃该动作")
            return False
        if not self.is_paused():
            self.pause(assume_running=True)

        if self.bullet_state() is not True:
            if not self._enter_bullet_chain(action):
                return False
        blind_enabled = self._axis_setting_bool("低回费盲步进", True)
        guard = 0
        while self.current_frame() < target:
            self._raise_if_stop()
            if self.bullet_state() is not True:
                self._dbg("  bullet 丢失，重新进子弹")
                if not self._enter_bullet_chain(action):
                    return False
            cur = self.current_frame()
            # 计数器可能在步进途中回零（离开战斗/重开）：进场时还正常，
            # 走到一半就变 0，于是会从 0 一路步进到上千帧。这里每步都查。
            if cur + FRAME_ROLLBACK_LIMIT < self._max_frame_seen:
                print(f"  WARNING: 步进中帧号从 {self._max_frame_seen} 回退到 "
                      f"{cur}，可能已离开战斗/重开；停止本次轴，请重新 run")
                self.stop_requested.set()
                return False
            self._max_frame_seen = max(self._max_frame_seen, cur)
            hi = self.current_frame_hi()
            # 只有目标落在“费用条同宽”的已知区间内才盲走；超过上界
            # 说明条还没恢复变化，盲走没有依据。
            allow_blind = bool(blind_enabled and cur < target <= hi)
            self.phase = f"步进中 {cur}→{target}"
            self.step_one_frame(require_bullet=True, allow_blind=allow_blind,
                                allow_overshoot=True)
            guard += 1
            if self.current_frame() > target:
                print(f"  WARNING: 过帧 target={target} "
                      f"current={self.current_frame()}")
                return False
        return self.current_frame() == target

    def _drag_deploy_checked(self, card, middle, drop, start_frame: int) -> str:
        """拖卡到目标格；返回 ok / invalid / f6。"""
        self.deploy_active = True
        try:
            self.touch.touch_down(*card)
            t0 = time.perf_counter()
            if self._wait_drag_or_f6(0.05, card):
                return "f6"
            self.touch.touch_move(*middle)
            if self._wait_drag_or_f6(DRAG_MIDDLE_WAIT, middle):
                return "f6"
            for i in range(1, DEPLOY_DRAG_STEPS + 1):
                p = (
                    int(middle[0] + (drop[0] - middle[0]) * i / DEPLOY_DRAG_STEPS),
                    int(middle[1] + (drop[1] - middle[1]) * i / DEPLOY_DRAG_STEPS),
                )
                self.touch.touch_move(*p)
                if self._wait_drag_or_f6(DEPLOY_DRAG_STEP, p):
                    return "f6"
                if (time.perf_counter() - t0 >= 0.1
                        and self.bullet_state() is not True):
                    try:
                        self.touch.touch_up(*p)
                    except Exception:
                        pass
                    return "invalid"
            if (time.perf_counter() - t0 >= 0.1
                    and self.bullet_state() is not True):
                try:
                    self.touch.touch_up(*drop)
                except Exception:
                    pass
                return "invalid"
            # 落点先停一下再抬手。拖拽每步 12ms，而游戏 60fps（16.7ms 一帧），
            # 手指移动和抬起贴得太近时，游戏可能还没处理到最后一次移动就
            # 收到抬起，按前一个中间点判定落点 —— 格子靠边时就会
            # “拖过去了却没放上”（2026-11-09 1110 帧 F9 就是这种）。
            time.sleep(DROP_HOLD)
            self.touch.touch_up(*drop)
            return "ok"
        finally:
            self.deploy_active = False

    def _drag_direction_checked(self, tile, action: AxisAction) -> bool:
        dx = dy = 0
        if action.direction == "右":
            dx = int(DIRECTION_DX * self.cw)
        elif action.direction == "左":
            dx = -int(DIRECTION_DX * self.cw)
        elif action.direction == "上":
            dy = -int(DIRECTION_DY * self.ch)
        elif action.direction == "下":
            dy = int(DIRECTION_DY * self.ch)
        if not (dx or dy):
            return False
        if self._drag_direction(tile, dx, dy):
            print("  F6：方向拖拽中断")
            self._cancel_bullet_time()
            self._manual_hold()
            return True
        time.sleep(0.15)
        return False

    def _target_card_present(self, oper: str) -> bool | None:
        """目标卡是否还在待部署区；None=这一帧没抓到/识别异常。"""
        frame = self.stream.latest()
        if frame is None:
            return None
        try:
            locate_avatar(frame, oper)
            return True
        except RuntimeError:
            return False
        except Exception as exc:
            print(f"  WARNING: 部署后检查目标卡失败：{exc!r}")
            return None

    def _card_count(self) -> int | None:
        """待部署区能数出几张卡（按旗标）。"""
        frame = self.stream.latest()
        if frame is None:
            return None
        try:
            return len(operator_cards(frame))
        except Exception:
            return None

    def _card_gone_stable(self, oper: str, misses: int = 3,
                          timeout: float = 0.9) -> bool:
        """连续 `misses` 次都认不出目标卡，才认为它真的没了。

        单次 `locate_avatar` 会偶发认不出（卡片行滚动、朝向界面变暗、
        卡片被选中的瞬间），一次 False 就下结论会把“没放上去”判成成功
        （30 帧那次就是）。
        """
        deadline = time.perf_counter() + timeout
        streak = 0
        while True:
            present = self._target_card_present(oper)
            if present is False:
                streak += 1
                if streak >= misses:
                    return True
            else:
                streak = 0
            if time.perf_counter() >= deadline:
                return False
            time.sleep(0.12)

    def _wait_target_card_consumed(self, oper: str,
                                   timeout: float = 0.6) -> bool:
        """部署后确认目标卡已经从待部署区消失（连续多次认不出才算）。"""
        if self._card_gone_stable(oper, misses=3, timeout=timeout):
            print(f"  部署验证：{oper} 的待部署卡已消失")
            return True
        if self._target_card_present(oper) is True:
            print(f"  WARNING: 部署后 {oper} 的待部署卡仍在")
        else:
            print(f"  WARNING: 部署后无法确认 {oper} 的待部署卡状态")
        return False

    def _cancel_pending_deploy(self) -> None:
        """点掉卡在“选择朝向”界面上的半成品部署。

        干员放下了但朝向没设成时，游戏会一直停在朝向界面并显示
        “点击取消”（屏幕上方偏右的红叉）。不点掉它，后面每一次点击
        都被这个界面吃掉：进不了子弹、点卡片也点不准。
        不在这个状态时这一下落在顶部 HUD 上，没有副作用。
        """
        self.touch.tap(*self.screen(0.636, 0.045))
        time.sleep(0.18)
        if self.bullet_state() is True:
            # 取消后可能回到“卡选中”态，再点一次退出。
            self.touch.tap(*self.screen(0.636, 0.045))
            time.sleep(0.15)
        print("  已点击“取消”退出未完成的部署界面")

    def _save_deploy_fail(self, frame: int, oper: str | None,
                          drop: tuple[int, int]) -> None:
        """部署失败时存一张带落点标记的截图，方便核对落点是否压在格子上。

        落点对不对光看日志坐标判断不了，图上一画就清楚。
        """
        try:
            import cv2
            img = self.stream.latest()
            if img is None:
                return
            x = int(drop[0] - self.client_left)
            y = int(drop[1] - self.client_top)
            cv2.drawMarker(img, (x, y), (0, 0, 255),
                           markerType=cv2.MARKER_CROSS, markerSize=24,
                           thickness=2)
            os.makedirs("out", exist_ok=True)
            safe = "".join(ch if ch.isascii() and ch.isalnum() else "_"
                           for ch in str(oper or "x"))
            path = os.path.join("out", f"deploy_fail_{frame}_{safe}.png")
            # cv2.imwrite 在 Windows 上遇到中文路径会**静默返回 False**，
            # 之前日志说“已保存”其实什么都没写。改用 imencode + tofile。
            ok, buf = cv2.imencode(".png", img)
            if not ok:
                print("  WARNING: 落点截图编码失败")
                return
            buf.tofile(path)
            print(f"  已保存落点截图：{path}")
        except Exception as exc:
            print(f"  保存落点截图失败：{exc!r}")

    def _deploy_at_frame(self, action: AxisAction, frame: int) -> tuple[bool, int | None]:
        """已停在该帧的子弹时间；选目标卡、拖动、判定、调方向。"""
        self.phase = "部署：选择待部署卡"
        # 落点用 side（prts-plus 的约定，实机一直用它）。front 是干员真正站
        # 的位置，但那是“选中干员”用的点，不是落点；只在拖朝向时用。
        # 之前拿 front 当备用落点重试过，结果 D2 那种本来就该失败的行会
        # 连放两次，白花时间，已去掉。
        if action.view_side is None:
            print("  WARNING: 该行没有落点坐标，跳过部署")
            return False, None
        view = action.view_side
        while True:
            if self.manual_pause.is_set():
                self._manual_hold()
            try:
                img = self.stream.latest()
                ax, ay, score = locate_avatar(img, action.oper)
            except Exception as e:
                print(f"  WARNING: 找不到待部署卡: {e!r}")
                return False, None
            card = self.screen(ax, ay)
            self._dbg(f"  avatar {action.oper} at ({ax:.3f},{ay:.3f}) "
                      f"score={score:.3f} frame={self.current_frame()}")
            # 进子弹时点的就是最右卡；如果目标正好是这张卡，它已经
            # 选中了，再点一次会取消选中并退出子弹。
            # 进子弹现在点的是固定点，和目标卡头像中心天然差几十像素，
            # 容差要放宽到半个卡位以内（一个卡位约 0.09 屏宽）。
            entry = self._bullet_entry_card_screen
            already_selected = bool(
                entry is not None
                and abs(card[0] - entry[0]) <= int(0.05 * self.cw)
                and abs(card[1] - entry[1]) <= int(0.06 * self.ch))
            if already_selected:
                self._dbg("  目标卡已是进入子弹时选中的卡，跳过重复点击")
                time.sleep(0.05)
            else:
                self.touch.click(*card)
                time.sleep(CARD_CLICK_WAIT)
            self._bullet_entry_card_screen = None
            self.selected_oper = action.oper
            if self.bullet_state() is not True:
                if not self._enter_bullet_chain(action):
                    return False, None
            # 选中后待部署区会滚动，重新找卡。
            try:
                img = self.stream.latest()
                ax, ay, _ = locate_avatar(img, action.oper)
                card = self.screen(ax, ay)
            except Exception:
                pass

            tile = self.screen(*view)
            drop = (tile[0], tile[1] + int(0.02 * self.ch))
            middle = (card[0], card[1] - int(0.03 * self.ch))
            self._dbg(f"  deploy drop=({drop[0]},{drop[1]}) "
                      f"from side=({view[0]:.3f},{view[1]:.3f})")
            self.phase = "部署：拖动中"
            res = self._drag_deploy_checked(card, middle, drop, frame)
            if res == "f6":
                self._manual_hold()
                if self.current_frame() == frame:
                    print("  F6 恢复后仍在指定帧，重新部署")
                    continue
                print("  WARNING: F6 恢复后已不在指定帧")
                return False, None
            if res == "invalid":
                print("  不可部署：拖动途中丢失子弹时间，放弃本次部署")
                self._exit_bullet_chain()
                return False, None

            self.phase = "部署：落点判定"
            time.sleep(DROP_WAIT)
            # 松手后游戏会立刻退出子弹时间并弹出“选择朝向”的四个箭头。
            # 也就是说 bullet=False 才是**放置成功**的正常表现，不能当失败判据。
            # 2026-11-09 实测：F9 明明已经放上去、朝向箭头都出来了，却因为
            # bullet=False 被判失败 → 游戏停在朝向界面 → 后面克洛丝进不了子弹、
            # 装置技能点到右数第二张卡。成败一律交给“待部署卡是否被消耗”。
            if self.bullet_state() is not True:
                print("  落点后已退出子弹时间（弹出朝向箭头属正常）")

            self.phase = "部署：调整方向"
            # 严格按轴：从轴给的落点拖一次朝向，不重试、不换锚点。
            interrupted = False
            if action.direction in ("上", "下", "左", "右"):
                if self._drag_direction_checked(tile, action):
                    interrupted = True      # F6 打断，交给外层决定是否重来
            else:
                # 没填朝向：等游戏自己把朝向界面收掉。
                time.sleep(0.3)
            if interrupted:
                if self.current_frame() != frame:
                    return False, None
                continue

            actual = self.current_frame()
            if not self._settle_paused():
                print("  WARNING: 部署后退出子弹时间失败，判定部署未成功")
                print("  已停止后续动作；请先手动退出子弹，再重新 run")
                self.stop_requested.set()
                if self.reporter:
                    self.reporter.set_error(
                        "部署后退出子弹时间失败；已停止后续动作")
                return False, None

            # 必须先退出子弹再验证。卡还在手里（bullet=True）时，它的旗标
            # 本来就不存在，用“旗标找不到”判断“卡已消失”必然假阳性——
            # 30 帧那次费用不够、卡根本没放下去，却被判成部署成功。
            if action.oper and not self._wait_target_card_consumed(action.oper):
                self._save_deploy_fail(frame, action.oper, drop)
                # 卡还在说明这半次放置没完成；如果游戏停在“选择朝向”界面，
                # 必须先点掉“点击取消”，否则后面所有点击都会被它吃掉
                # （2026-11-09：芬 卡在朝向界面 → 克洛丝进不了子弹）。
                self._cancel_pending_deploy()
                print("  WARNING: 部署后目标卡仍在待部署区，判定部署未成功")
                if self.reporter:
                    self.reporter.set_error(
                        "部署后目标卡仍在待部署区；判定部署未成功")
                return False, None
            # 记录实际落点，供后续进子弹/撤退使用。
            self.deployed_stack = [
                r for r in self.deployed_stack
                if r.get("name") != action.oper
            ]
            self.deployed_stack.append({
                "name": action.oper,
                "alias": action.alias,
                "view_front": action.view_front,
                "view_side": action.view_side,
                "tile_pos": action.tile_pos,
                "direction": action.direction,
            })
            print(f"  deployed {action.oper} frame={actual}")
            return True, actual

    def _target_tile_check(self, action: AxisAction) -> bool:
        """技能/撤退前核对：这个**干员**真的在那个格子上吗？

        实机踩过：撤退行的“面前坐标”填成了另一个干员的格子，而那个干员
        部署失败根本不在场，结果 Q 把站在那格上的**别人**撤了。

        只对干员生效（干员表里 char_ 开头的 ID）；装置/召唤物（token_ 开头，
        如风雪之眼、投递坐标）可以和干员同格，不做这套核对。
        """
        name = getattr(action, "oper", None)
        view = getattr(action, "view_front", None)
        if not name or view is None or not self.deployed_stack:
            return True
        try:
            from .vision import operator_mapping
            char_id = operator_mapping().get(name) or ""
        except Exception:
            char_id = ""
        if char_id and not char_id.startswith("char_"):
            # 装置/召唤物：不参与格子归属核对
            return True
        same_tile = [
            r for r in self.deployed_stack
            if r.get("view_front") is not None
            and abs(float(r["view_front"][0]) - float(view[0])) < 0.01
            and abs(float(r["view_front"][1]) - float(view[1])) < 0.01
        ]
        owners = {r.get("name") for r in same_tile if r.get("name")}
        if same_tile and name not in owners:
            print(f"  WARNING: 目标格 {view} 在部署栈里属于 "
                  f"{'/'.join(sorted(owners))}，不是 {name}；"
                  f"这一下会打到别人，已跳过本行")
            return False
        if not any(r.get("name") == name for r in self.deployed_stack):
            print(f"  WARNING: {name} 没有成功部署过（部署栈里没有），跳过本行")
            return False
        return True

    def _skill_at_frame(self, action: AxisAction, frame: int) -> tuple[bool, int | None]:
        """已停在该帧的子弹时间；退出子弹 → AFA 三连点选人 → E/Q。"""
        if not self._target_tile_check(action):
            if self.bullet_state() is True:
                self._exit_bullet_chain()
            return False, None
        if not self._exit_bullet_chain():
            return False, None
        if self.is_running_icon():
            self.pause(assume_running=True)
        self.phase = f"{action.kind}：选中干员"
        print(f"  select {action.oper} at frame={self.current_frame()}")
        try:
            # 用统一的选中重试：默认 hold，失败会换 direct/tap，
            # 第一轮失败后还会尝试点最右卡重新进子弹。
            self._select_operator(action, target_frame=frame)
        except StopRequested:
            raise
        except FatalReloadError:
            raise
        except Exception as exc:
            # 装置（单位列为空）没有名字可校验，front 点位偏一点就选不中，
            # 用 side 点位（落点那个投影）再补一次。
            if (action.oper is None and action.view_side is not None
                    and action.view_side != action.view_front):
                print(f"  {action.kind}：front 点位选不中（{exc}），"
                      f"改用 side=({action.view_side[0]:.3f},"
                      f"{action.view_side[1]:.3f}) 再试")
                saved_front = action.view_front
                action.view_front = action.view_side
                try:
                    self._select_operator(action, target_frame=frame)
                except StopRequested:
                    raise
                except FatalReloadError:
                    raise
                except Exception as exc2:
                    print(f"  WARNING: 选中干员失败，放弃本次{action.kind}：{exc2}")
                    # 别把“选中态/子弹时间”留给下一个动作：先退出子弹并确认暂停。
                    self._settle_paused()
                    return False, None
                finally:
                    action.view_front = saved_front
            else:
                print(f"  WARNING: 选中干员失败，放弃本次{action.kind}：{exc}")
                self._settle_paused()
                return False, None
        if self.current_frame() != frame:
            print(f"  WARNING: 选人后帧号变了 target={frame} "
                  f"current={self.current_frame()}")
            if self.current_frame() > frame:
                return False, None
        self.selected_oper = action.oper
        vk = RETREAT_KEY if action.kind == "撤退" else SKILL_KEY
        label = "撤退 Q" if action.kind == "撤退" else "技能 E"
        self.phase = f"{action.kind}：按键 {'Q' if action.kind == '撤退' else 'E'}"
        ok = self._send_action_verified(action, vk, label, action.kind)
        actual = self.current_frame()
        if self.bullet_state() is True:
            self._settle_paused()
        elif self.is_running_icon():
            self.pause(assume_running=True)
        if ok and action.kind == "撤退" and action.oper:
            self.deployed_stack = [
                r for r in self.deployed_stack
                if r.get("name") != action.oper
            ]
        return ok, (actual if ok else None)

    def _pause_at_frame(self, action: AxisAction, frame: int) -> tuple[bool, int | None]:
        if not self._exit_bullet_chain():
            return False, None
        if self.is_running_icon():
            self.pause(assume_running=True)
        cur = self.current_frame()
        if cur != frame or self.bullet_state() is True:
            print(f"  WARNING: 暂停落点不对 target={frame} current={cur} "
                  f"bullet={self.bullet_state()}")
            return False, None
        self.phase = "暂停操作：等待手动恢复"
        print(f"  暂停操作已停在 frame={cur}，等待手动恢复")
        self.wait_manual_resume()
        return True, frame

    def _attempt_action(self, action: AxisAction,
                        frame: int) -> tuple[bool, int | None, bool]:
        """返回 (成功, 实际帧, 是否因 J=False 跳过)。"""
        idx = getattr(action, "_index", -1)
        if self.reporter:
            self.reporter.action_start(idx, action)
        self.current_action = (idx, action)
        self.selected_oper = None
        try:
            if not self._prepare_to_frame(action, frame):
                self._settle_paused()
                return False, None, False
            # 游戏已停在目标帧，同步屏障读 J 最可靠。
            if not self._read_enabled_now(idx, action):
                self.phase = f"{action.kind}：是否执行=False，跳过"
                self._settle_paused()
                return False, None, True
            if action.kind == "部署":
                ok, actual = self._deploy_at_frame(action, frame)
            elif action.kind in ("技能", "撤退"):
                ok, actual = self._skill_at_frame(action, frame)
            elif action.kind == "暂停":
                ok, actual = self._pause_at_frame(action, frame)
            else:
                ok, actual = False, None
            if not ok:
                settled = self._settle_paused()
                if settled is False and not self.stop_requested.is_set():
                    print("  WARNING: 动作失败后退出子弹时间也失败；"
                          "停止后续动作")
                    self.stop_requested.set()
                    if self.reporter:
                        self.reporter.set_error(
                            "动作失败后退出子弹时间失败；已停止后续动作")
            return ok, actual, False
        except FatalReloadError:
            raise
        except StopRequested:
            raise
        except Exception as exc:
            row = getattr(action, "excel_row", 0) or idx
            msg = f"第 {row} 行（{action.kind}）执行异常：{exc}"
            print(f"  [错误] {msg}")
            if self.reporter:
                self.reporter.set_error(f"错误：{msg}")
            self._settle_paused()
            return False, None, False
        finally:
            self.current_action = None
            self.selected_oper = None
            if self.phase != "等待修复":
                self.phase = "运行中"

    def _read_enabled_now(self, idx: int, action: AxisAction) -> bool:
        rep = self.reporter
        # 游戏在跑时不去等 Excel，避免过帧；等停到目标帧后再同步读。
        if rep is not None and hasattr(rep, "sync_read_enabled") \
                and (self.is_paused() or self.bullet_state() is True):
            val = rep.sync_read_enabled(idx)
            if val is not None:
                action.enabled = bool(val)
                return bool(val)
        return bool(getattr(action, "enabled", True))

    def _commit_result_now(self, idx: int, ok: bool | None) -> None:
        rep = self.reporter
        if rep is not None and hasattr(rep, "sync_commit_result"):
            rep.sync_commit_result(idx, ok)

    def _finish_action(self, idx: int, action: AxisAction,
                       actual: int | None, ok: bool,
                       error: str = "") -> None:
        if self.reporter:
            note = "成功" if ok else "失败"
            if action.kind == "无操作":
                note = "无操作"
            self.reporter.action_done(idx, action, action.frame_text,
                                      actual, note)
            if error:
                self.reporter.set_error(error)

    @staticmethod
    def _build_rate_events(selected):
        """把“回费速度”行展开成按帧生效的事件。

        返回 (events, last_frame)：
          events = [(帧号, 手速百分比或None, 对应动作或None)]
          None 表示恢复设置里的默认回费速度。
        last_frame 为最后一个事件帧；没有事件时为 None。
        """
        events = []
        last = None
        for k, action in enumerate(selected):
            if action.kind != "回费速度" or not getattr(action, "enabled", True):
                continue
            idx = getattr(action, "_index", k)
            value = None if action.rate_auto else action.rate_percent
            if value is not None and not (0.0 < float(value) <= 1000.0):
                continue
            row = int(getattr(action, "excel_row", 0) or 0)
            frame = int(action.frame)
            events.append((frame, 0, row, idx, value, action))
            last = frame if last is None else max(last, frame)
            if action.frame2 is not None:
                # C 是最后一帧：C+1 恢复默认。
                end = int(action.frame2) + 1
                events.append((end, -1, row, idx, None, None))
                last = end if last is None else max(last, end)
        events.sort(key=lambda e: (e[0], e[1], e[2], e[3]))
        # 同一个事件帧上，结束(-1)先于开始(0)，这样新一段从 C+1 开始能覆盖默认。
        return [(e[0], e[4], e[5]) for e in events], last

    def _rate_action_applied(self, action) -> None:
        """计数线程在回费速度事件生效时回调，记录该行“成功”。"""
        rep = self.reporter
        if rep is None:
            return
        idx = getattr(action, "_index", None)
        if idx is None:
            return
        if hasattr(rep, "record_simple_action"):
            rep.record_simple_action(idx, action, action.frame, "成功")
        else:
            rep.action_done(idx, action, action.frame_text,
                            action.frame, "成功")

    @staticmethod
    def _apply_stack_entry(stack: list[dict], action) -> None:
        if not getattr(action, "enabled", True):
            return
        if action.kind == "部署" and action.oper:
            stack[:] = [r for r in stack if r.get("name") != action.oper]
            stack.append({
                "name": action.oper, "alias": action.alias,
                "view_front": action.view_front,
                "view_side": action.view_side,
                "tile_pos": action.tile_pos,
                "direction": action.direction,
            })
        elif action.kind == "撤退" and action.oper:
            stack[:] = [r for r in stack if r.get("name") != action.oper]

    def _rebuild_stack_for_start(self, initial_cur: int,
                                 start_action: int, selected) -> None:
        """按轴历史重建部署栈：start_action 之前的 + 已过帧的选中动作。"""
        stack = list(self.deployed_stack)
        for a in self.axis.actions[:start_action]:
            self._apply_stack_entry(stack, a)
        for a in selected:
            if a.end_frame < initial_cur:
                self._apply_stack_entry(stack, a)
        self.deployed_stack = stack

    @staticmethod
    def _trim_rate_events(events, rate_last, initial_cur: int):
        """去掉当前帧之前的回费事件，但保留当前生效的手动回费速度。"""
        if initial_cur <= 0 or not events:
            return events, rate_last
        active = None
        kept = []
        for frame, value, action in events:
            if frame < initial_cur:
                active = value
            else:
                kept.append((frame, value, action))
        if active is not None:
            # action=None：这是“当前已经生效”的状态，不把它算作执行成功。
            kept.insert(0, (initial_cur, active, None))
            kept.sort(key=lambda e: e[0])
        new_last = max((e[0] for e in kept), default=None)
        return kept, new_last

    def run(self, max_actions: int = 1, start_action: int = 0) -> None:
        """按帧推进：同帧内按 (帧号1, 表格行号) 依次尝试。"""
        self._max_frame_seen = self.current_frame()
        selected = self.axis.actions[start_action:start_action + max_actions]
        initial_cur = self.current_frame()
        self._rebuild_stack_for_start(initial_cur, start_action, selected)
        pending: list[tuple[int, AxisAction]] = []
        skipped_rows: list[str] = []
        for k, action in enumerate(selected):
            idx = start_action + k
            action._index = idx
            if action.kind in ("无操作", "回费速度"):
                # 空操作不执行；回费速度走单独的按帧事件，不占动作队列。
                continue
            if action.end_frame < initial_cur:
                row = getattr(action, "excel_row", 0) or (idx + 1)
                skipped_rows.append(str(row))
                continue
            if action.kind not in ("部署", "技能", "撤退", "暂停"):
                row = getattr(action, "excel_row", 0) or idx
                self._finish_action(
                    idx, action, None, False,
                    f"错误：第 {row} 行操作类型“{action.kind}”无法识别；"
                    "支持：部署/技能/撤退/暂停/无操作/回费速度。"
                    "该行未执行。")
                continue
            pending.append((idx, action))
        if skipped_rows:
            print(f"  跳过 {len(skipped_rows)} 个已过帧动作："
                  f"第 {', '.join(skipped_rows)} 行")

        rate_events, rate_last = self._build_rate_events(selected)
        rate_events, rate_last = self._trim_rate_events(
            rate_events, rate_last, initial_cur)
        counter = getattr(self, "counter", None)
        if counter is not None and hasattr(counter, "set_rate_events"):
            counter.rate_callback = self._rate_action_applied
            counter.set_rate_events(rate_events)

        def _rate_done() -> bool:
            return (counter is None or not hasattr(counter, "rate_events_done")
                    or counter.rate_events_done())

        self.stop_requested.clear()
        stopped = False
        self.phase = "运行中"
        self.axis_running = True
        attempted: dict[int, set[int]] = {}
        while pending or (rate_last is not None and not _rate_done()):
            if self.stop_requested.is_set():
                stopped = True
                print("  停止请求：退出执行循环")
                break
            cur = self.current_frame()
            # 已经越过区间的操作：启用的写 FALSE，禁用/空操作静默丢弃。
            for idx, action in list(pending):
                if cur > action.end_frame:
                    if getattr(action, "enabled", True):
                        self._commit_result_now(idx, False)
                        row = getattr(action, "excel_row", 0) or idx
                        self._finish_action(
                            idx, action, None, False,
                            f"错误：第 {row} 行已过帧：目标 {action.frame}，"
                            f"当前 {cur}；该行未执行。")
                    pending.remove((idx, action))
            if not pending:
                if rate_last is None or _rate_done() or cur > rate_last:
                    break
                # 只剩回费速度事件：等计数线程按帧应用，不抢时序。
                if self.manual_pause.is_set():
                    self._manual_hold()
                time.sleep(0.01)
                continue

            # 本帧候选：不先用缓存 enabled 过滤，逐行同步读 J。
            due = [
                (idx, a) for idx, a in pending
                if a.frame <= cur <= a.end_frame
                and cur not in attempted.get(idx, set())
            ]
            due.sort(key=lambda x: (x[1].frame, x[1].excel_row, x[0]))
            chosen = None
            for idx, a in due:
                attempted.setdefault(idx, set()).add(cur)
                if self._read_enabled_now(idx, a):
                    chosen = (idx, a)
                    break
            if chosen is None:
                # 当前帧没有可执行行：挑下一个目标帧（跳过已试过的帧）。
                cands = []
                for idx, a in pending:
                    if not getattr(a, "enabled", True):
                        continue
                    t = max(cur + 1, a.frame)
                    tried = attempted.get(idx, set())
                    while t in tried and t <= a.end_frame:
                        t += 1
                    if t <= a.end_frame:
                        cands.append((t, idx, a))
                if not cands:
                    break
                target, idx, action = min(
                    cands, key=lambda x: (x[0], x[2].frame,
                                          x[2].excel_row, x[1]))
                attempted.setdefault(idx, set()).add(target)
                if not self._read_enabled_now(idx, action):
                    continue
                try:
                    ok, actual, skipped = self._attempt_action(action, target)
                except StopRequested:
                    stopped = True
                    print("  停止请求：从等待中退出")
                    break
                if skipped:
                    continue
                final = (not ok) and target >= action.end_frame
                self._commit_result_now(idx, True if ok else
                                        (False if final else None))
                if ok:
                    self._finish_action(idx, action, actual, True)
                    pending.remove((idx, action))
                elif final:
                    self._finish_action(idx, action, None, False)
                    pending.remove((idx, action))
                continue

            idx, action = chosen
            try:
                ok, actual, skipped = self._attempt_action(action, cur)
            except StopRequested:
                stopped = True
                print("  停止请求：从等待中退出")
                break
            if skipped:
                continue
            final = (not ok) and cur >= action.end_frame
            self._commit_result_now(idx, True if ok else
                                    (False if final else None))
            if ok:
                self._finish_action(idx, action, actual, True)
                pending.remove((idx, action))
            elif final:
                self._finish_action(idx, action, None, False)
                pending.remove((idx, action))

        self.axis_running = False
        self._write_pause_log()
        self.phase = "已停止" if stopped else "轴完成"
        if stopped:
            self.manual_pause.clear()
            if self.is_running_icon():
                if foreground_is(self.hwnd):
                    self.pause(assume_running=True)
                else:
                    print("  轴已停止：游戏不在前台，不切回游戏，也不点击暂停")
            print("  轴已停止")
            return
        if self.manual_pause.is_set():
            print("  轴已完成，但手动暂停仍生效：游戏保持暂停")
            try:
                self._manual_hold()
            except StopRequested:
                self.manual_pause.clear()
                print("  轴已停止")
                return
        if self.end_pause:
            if self.is_running_icon():
                self.pause()
            print("  轴完成：结束后暂停（轴设置），游戏保持暂停")
            if self.reporter is not None:
                if self.reporter.bring_to_front():
                    print("  轴完成：Excel 已前置")
                else:
                    print("  WARNING: 轴完成，但 Excel 没能前置")
        else:
            self.resume()
