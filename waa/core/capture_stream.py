"""持续 WGC 抓屏，保留最新一帧客户区画面。

逐帧推进时需要频繁读费用条，单次 WGC 启动/停止太慢（约 0.2s），
这里用常驻抓屏线程，读的时候直接拿最新帧。
"""

from __future__ import annotations

import queue
import threading
import time

import numpy as np

from windows_capture import WindowsCapture

from .capture import client_offset, client_rect


class CaptureStream:
    def __init__(self, hwnd: int):
        self.hwnd = hwnd
        self.ox, self.oy = client_offset(hwnd)
        _, _, self.cw, self.ch = client_rect(hwnd)
        self._frame: np.ndarray | None = None
        self._seq = 0
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        # 按到达顺序排队，计数线程逐帧消费；不能只取 latest，否则会丢帧。
        self._queue: queue.Queue[np.ndarray] = queue.Queue(maxsize=120)

    def start(self, timeout: float = 3.0) -> None:
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        deadline = time.time() + timeout
        while time.time() < deadline:
            if self.latest() is not None:
                return
            time.sleep(0.02)
        raise TimeoutError("WGC 持续抓屏没有拿到第一帧")

    def _run(self) -> None:
        capture = WindowsCapture(window_hwnd=self.hwnd, cursor_capture=False, draw_border=False)

        @capture.event
        def on_frame_arrived(frame, capture_control):
            if self._stop.is_set():
                capture_control.stop()
                return
            try:
                buf = frame.frame_buffer
                x1 = min(buf.shape[1], self.ox + self.cw)
                y1 = min(buf.shape[0], self.oy + self.ch)
                cropped = buf[self.oy:y1, self.ox:x1, :3].copy()
            except Exception:
                return
            with self._lock:
                self._frame = cropped
                self._seq += 1
            try:
                self._queue.put_nowait(cropped)
            except queue.Full:
                # 消费端跟不上时丢最旧的，别阻塞 WGC 回调
                try:
                    self._queue.get_nowait()
                except queue.Empty:
                    pass
                try:
                    self._queue.put_nowait(cropped)
                except queue.Full:
                    pass

        @capture.event
        def on_closed():
            self._stop.set()

        capture.start()

    def latest(self) -> np.ndarray | None:
        with self._lock:
            return self._frame

    def latest_seq(self) -> int:
        with self._lock:
            return self._seq

    def get_frame(self, timeout: float = 0.02) -> np.ndarray | None:
        """按到达顺序取一帧；超时返回 None。"""
        try:
            return self._queue.get(timeout=timeout)
        except queue.Empty:
            return None

    def stop(self) -> None:
        self._stop.set()
        # WGC 在下一帧到来时会自己停；这里不强制 join，避免阻塞退出
