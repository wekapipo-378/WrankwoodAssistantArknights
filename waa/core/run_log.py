"""把一次 run 的完整控制台输出同时写到独立日志文件。

不修改打印调用：临时替换 sys.stdout/sys.stderr 为 tee，所有线程的
print（主线程、执行器、Excel 面板线程）都会进入同一个文件。
"""

from __future__ import annotations

import os
import sys
import threading
import time


class _Tee:
    def __init__(self, stream, file_obj, lock: threading.Lock):
        self._stream = stream
        self._file = file_obj
        self._lock = lock

    def write(self, data):
        if not isinstance(data, str):
            data = str(data)
        with self._lock:
            if self._stream is not None:
                try:
                    self._stream.write(data)
                except Exception:
                    pass
            try:
                self._file.write(data)
            except Exception:
                pass
        return len(data)

    def flush(self):
        with self._lock:
            if self._stream is not None:
                try:
                    self._stream.flush()
                except Exception:
                    pass
            try:
                self._file.flush()
            except Exception:
                pass

    def __getattr__(self, name):
        return getattr(self._stream, name)


class RunLogger:
    def __init__(self, path: str, header: str = ""):
        self.path = path
        self.header = header
        self._file = None
        self._lock = threading.Lock()
        self._old_stdout = None
        self._old_stderr = None
        self._active = False

    def start(self) -> None:
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        self._file = open(self.path, "a", encoding="utf-8", buffering=1)
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        # 版本行：排查时先确认是哪个版本跑出来的。延迟 import 避免包初始化顺序问题。
        try:
            from waa import VERSION_CHANNEL, __version__
            ver = f"{__version__}（{VERSION_CHANNEL}）"
        except Exception:
            ver = "unknown"
        self._file.write(
            "=" * 78 + "\n"
            f"WAA run log  started={ts}\n"
            f"version={ver}\n"
            f"log={self.path}\n"
        )
        if self.header:
            self._file.write(self.header.rstrip() + "\n")
        self._file.write("=" * 78 + "\n")
        self._file.flush()
        self._old_stdout = sys.stdout
        self._old_stderr = sys.stderr
        sys.stdout = _Tee(self._old_stdout, self._file, self._lock)
        sys.stderr = _Tee(self._old_stderr, self._file, self._lock)
        self._active = True

    def command(self, line: str) -> None:
        """记录用户在控制台输入的命令（终端本身不会经过 print）。"""
        if not self._active or self._file is None:
            return
        with self._lock:
            try:
                self._file.write(f"waa> {line}\n")
                self._file.flush()
            except Exception:
                pass

    def stop(self, footer: str = "") -> None:
        if not self._active:
            return
        self._active = False
        # 先恢复真实输出，再写尾注，避免尾注再次经过 tee。
        try:
            sys.stdout = self._old_stdout
            sys.stderr = self._old_stderr
        except Exception:
            pass
        with self._lock:
            try:
                if self._file is not None:
                    ts = time.strftime("%Y-%m-%d %H:%M:%S")
                    self._file.write("=" * 78 + "\n")
                    if footer:
                        self._file.write(footer.rstrip() + "\n")
                    self._file.write(f"finished={ts}\n")
                    self._file.write("=" * 78 + "\n")
                    self._file.flush()
                    self._file.close()
            except Exception:
                pass
            finally:
                self._file = None
