"""运行路径：区分「可写工作目录」和「只读打包资源」。

开发时两者都是项目根；PyInstaller 打包后：
  - app_root()     = exe 所在目录（out/、axes/、日志都写这里）
  - bundle_root()  = sys._MEIPASS（waa/resources、waa/templates 随包解压在这里）

所有需要读写用户数据的地方用 app_root()，只读的随包资源用 bundle_root()。
"""

from __future__ import annotations

import os
import sys


def is_frozen() -> bool:
    """是否运行在 PyInstaller 打出的 exe 里。"""
    return bool(getattr(sys, "frozen", False))


def app_root() -> str:
    """可写工作目录（exe 同级目录；开发时是项目根）。"""
    if is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def bundle_root() -> str:
    """只读资源根（打包后是 _MEIPASS，否则同 app_root）。"""
    return getattr(sys, "_MEIPASS", None) or app_root()


def asset_path(*parts: str) -> str:
    """拼一个随包资源路径。"""
    return os.path.join(bundle_root(), *parts)
