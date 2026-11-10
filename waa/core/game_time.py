"""读取当前游戏时间：费用数字 + 费用条相位。

坐标沿用 prts-plus 的 1280x720 比例（用户已确认只支持 16:9、UI 缩放 100）。
费用条像素用 prts-plus 的算法：整块费用区域二值化，取最后一行，白色占比 × 29。
费用数字用 Windows 自带 OCR（不需要 tesseract 模型），并按图像内容缓存。
"""

from __future__ import annotations

import asyncio
import hashlib
import os
import tempfile
from dataclasses import dataclass

import cv2
import numpy as np

from winrt.windows.globalization import Language
from winrt.windows.graphics.imaging import BitmapDecoder
from winrt.windows.media.ocr import OcrEngine
from winrt.windows.storage import FileAccessMode, StorageFile

SCREEN_W = 1280
SCREEN_H = 720
TICK_MAX = 30
WHITE_THRESHOLD = 160

# prts-plus GameRatioConfig
COST_AREA_RATIO = (0.906, 0.685, 1.0, 0.755)
# 0.33 会切掉两位数的十位（实测 "10" 只剩 "0"）；0.22 能完整包含数字，
# 只会带上费用图标右缘，OCR 读成 ">" 之类的非数字，过滤掉即可。
COST_NUMBER_AREA_RATIO = (0.22, 0.0, 1.0, 0.96)


@dataclass(frozen=True, order=True)
class GameTime:
    cost: int
    tick: int

    @property
    def total_ticks(self) -> int:
        return self.cost * TICK_MAX + self.tick


def _resize_to_standard(frame_bgr: np.ndarray) -> np.ndarray:
    h, w = frame_bgr.shape[:2]
    if (w, h) != (SCREEN_W, SCREEN_H):
        frame_bgr = cv2.resize(frame_bgr, (SCREEN_W, SCREEN_H), interpolation=cv2.INTER_AREA)
    return frame_bgr


def cost_area(frame_bgr: np.ndarray) -> np.ndarray:
    img = _resize_to_standard(frame_bgr)
    l, t, r, b = COST_AREA_RATIO
    return img[int(t * SCREEN_H):int(b * SCREEN_H), int(l * SCREEN_W):int(r * SCREEN_W)]


def read_tick(frame_bgr: np.ndarray) -> int:
    """费用条最后一行白色占比 → 0..29。"""
    area = cost_area(frame_bgr)
    gray = cv2.cvtColor(area, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, WHITE_THRESHOLD, 255, cv2.THRESH_BINARY)
    last_row = binary[-1, :]
    ratio = float(np.mean(last_row == 255))
    return round(ratio * (TICK_MAX - 1))


def _crop_cost_number(frame_bgr: np.ndarray) -> np.ndarray:
    area = cost_area(frame_bgr)
    h, w = area.shape[:2]
    l, t, r, b = COST_NUMBER_AREA_RATIO
    return area[int(t * h):int(b * h), int(l * w):int(r * w)]


# ---- Windows OCR ----
_ocr_engine = OcrEngine.try_create_from_language(Language("en-US"))
_cost_cache: dict[str, int] = {}


async def _ocr_png(path: str) -> str:
    file = await StorageFile.get_file_from_path_async(path)
    stream = await file.open_async(FileAccessMode.READ)
    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()
    if _ocr_engine is None:
        raise RuntimeError("没有可用的 Windows OCR 语言包")
    result = await _ocr_engine.recognize_async(bitmap)
    return result.text


def read_cost(frame_bgr: np.ndarray) -> int:
    """费用数字 OCR；同一张图只识别一次。识别失败返回 -1。"""
    crop = _crop_cost_number(frame_bgr)
    key = hashlib.md5(crop.tobytes()).hexdigest()
    if key in _cost_cache:
        return _cost_cache[key]

    # 实测放大后 Windows OCR 反而读不出（4x 三次插值返回空），原尺寸最稳
    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        cv2.imwrite(path, crop)
        text = asyncio.run(_ocr_png(path))
    finally:
        try:
            os.remove(path)
        except OSError:
            pass

    digits = "".join(ch for ch in text if ch.isdigit())
    cost = int(digits) if digits else -1
    _cost_cache[key] = cost
    # 缓存别无限涨
    if len(_cost_cache) > 200:
        _cost_cache.clear()
    return cost


def read_game_time(frame_bgr: np.ndarray) -> GameTime:
    return GameTime(cost=read_cost(frame_bgr), tick=read_tick(frame_bgr))
