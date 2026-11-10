"""Windows 自带 OCR：读取屏幕左侧当前选中干员的名字。

只裁左上角固定区域，后台线程识别，避免阻塞打轴主循环。
需要在 .venv 里安装 winrt-runtime 及 Windows.Media.Ocr 等 PyWinRT 包。
"""

from __future__ import annotations

import asyncio
import io
import re
import threading

import numpy as np


def _ocr_png(data: bytes) -> str:
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.globalization import Language
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream

    async def run() -> str:
        stream = InMemoryRandomAccessStream()
        writer = DataWriter(stream)
        writer.write_bytes(data)
        await writer.store_async()
        await writer.flush_async()
        stream.seek(0)
        decoder = await BitmapDecoder.create_async(stream)
        bitmap = await decoder.get_software_bitmap_async()
        engine = (OcrEngine.try_create_from_language(Language("zh-Hans-CN"))
                  or OcrEngine.try_create_from_user_profile_languages())
        if engine is None:
            return ""
        result = await engine.recognize_async(bitmap)
        best = ""
        for line in result.lines:
            text = re.sub(r"\s+", "", line.text or "")
            # 只保留中文、字母、数字和连字符；去掉等级/纯数字杂讯。
            text = "".join(re.findall(r"[一-鿿A-Za-z0-9\-]", text))
            if len(text) < 2:
                continue
            if text.lower() in ("lv", "tacx") or text.isdigit():
                continue
            if len(text) > len(best):
                best = text
        return best

    return asyncio.run(run())


def _ocr_lines(data: bytes) -> list[str]:
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.globalization import Language
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream

    async def run() -> list[str]:
        stream = InMemoryRandomAccessStream()
        writer = DataWriter(stream)
        writer.write_bytes(data)
        await writer.store_async()
        await writer.flush_async()
        stream.seek(0)
        decoder = await BitmapDecoder.create_async(stream)
        bitmap = await decoder.get_software_bitmap_async()
        engine = (OcrEngine.try_create_from_language(Language("zh-Hans-CN"))
                  or OcrEngine.try_create_from_user_profile_languages())
        if engine is None:
            return []
        result = await engine.recognize_async(bitmap)
        return [line.text or "" for line in result.lines]

    return asyncio.run(run())


def ocr_text_boxes(image_bgr, roi=None) -> list[dict]:
    """OCR 一个 BGR 图像/ROI，返回带像素外接框的行。

    返回项：{"text": 原文, "x","y","w","h": 在原图中的像素坐标}。
    用于找“开始行动/开始”按钮，不能只依赖固定模板。
    """
    from PIL import Image

    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.globalization import Language
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream

    x0 = y0 = 0
    if roi is not None:
        x0, y0, rw, rh = (int(v) for v in roi)
        image_bgr = image_bgr[y0:y0 + rh, x0:x0 + rw]
    if image_bgr is None or image_bgr.size == 0:
        return []
    im = Image.fromarray(image_bgr[:, :, ::-1])
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    data = buf.getvalue()

    async def run() -> list[dict]:
        stream = InMemoryRandomAccessStream()
        writer = DataWriter(stream)
        writer.write_bytes(data)
        await writer.store_async()
        await writer.flush_async()
        stream.seek(0)
        decoder = await BitmapDecoder.create_async(stream)
        bitmap = await decoder.get_software_bitmap_async()
        engine = (OcrEngine.try_create_from_language(Language("zh-Hans-CN"))
                  or OcrEngine.try_create_from_user_profile_languages())
        if engine is None:
            return []
        result = await engine.recognize_async(bitmap)
        out = []
        for line in result.lines:
            words = []
            boxes = []
            for word in line.words:
                r = word.bounding_rect
                boxes.append((r.x, r.y, r.width, r.height))
                words.append({
                    "text": word.text or "",
                    "x": int(r.x) + x0,
                    "y": int(r.y) + y0,
                    "w": int(r.width),
                    "h": int(r.height),
                })
            if not boxes:
                continue
            left = min(b[0] for b in boxes)
            top = min(b[1] for b in boxes)
            right = max(b[0] + b[2] for b in boxes)
            bottom = max(b[1] + b[3] for b in boxes)
            out.append({
                "text": line.text or "",
                "x": int(left) + x0,
                "y": int(top) + y0,
                "w": int(right - left),
                "h": int(bottom - top),
                "words": words,
            })
        return out

    try:
        return asyncio.run(run())
    except Exception:
        return []


def _norm(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def _parse_hud(left_lines: list[str], right_lines: list[str]) -> dict:
    left = [_norm(x) for x in left_lines if x]
    right = [_norm(x) for x in right_lines if x]
    out = {"name": None, "attack": None, "defense": None,
           "resist": None, "block": None, "hp_cur": None, "hp_max": None,
           "cost": None, "deploy_left": None}

    bad = ("攻击", "防御", "法抗", "阻挡", "生命", "技能", "特性",
           "天赋", "等级", "精英", "自动", "手动", "部署", "费用")
    for line in left:
        if 2 <= len(line) <= 8 and not any(k in line for k in bad) \
                and all("一" <= ch <= "鿿" or ch.isalpha()
                        for ch in line):
            out["name"] = line
            break

    joined = "".join(left)
    for key, label in (("attack", "攻击"), ("defense", "防御"),
                       ("resist", "法抗"), ("block", "阻挡")):
        m = re.search(label + r"\D{0,8}?(\d{1,5})", joined)
        if m:
            out[key] = int(m.group(1))

    hp = re.findall(r"(\d{1,6})/(\d{1,6})", joined)
    if hp:
        cur, mx = max(hp, key=lambda x: int(x[1]))
        out["hp_cur"], out["hp_max"] = int(cur), int(mx)

    for line in right:
        m = re.search(r"剩余可放置角色[：:]*(\d{1,3})", line)
        if m:
            out["deploy_left"] = int(m.group(1))
            break
    # 费用可能是负数（局内负费）。OCR 的减号可能是 ASCII '-'、全角 '－'、
    # 数学负号 '−' 或破折号；统一后再解析，别把负号丢掉。
    pending_minus = False
    for line in right:
        if "剩余" in line or "MAX" in line.upper():
            continue
        line = re.sub(r"[−－–—]", "-", line)
        if line and set(line) == {"-"}:
            # OCR 把减号单独识别成一行时，记下来给下一行数字用。
            pending_minus = True
            continue
        m = re.search(r"(-?)(\d{1,5})", line)
        if m:
            value = int(m.group(2))
            if m.group(1) or pending_minus:
                value = -value
            out["cost"] = value
            break
    return out


class HudOCR:
    """后台 OCR：左侧干员面板 + 右下角费用/剩余可放置。"""

    def __init__(self):
        self._lock = threading.Lock()
        self._pending = None
        self._values: dict = {}
        self._event = threading.Event()
        self._stop = False
        self._thread: threading.Thread | None = None

    def submit(self, image) -> None:
        try:
            from PIL import Image

            h, w = image.shape[:2]
            if image.ndim == 3 and image.shape[2] >= 3:
                image = np.ascontiguousarray(image[:, :, 2::-1])
            im = Image.fromarray(image)
            left = im.crop((0, int(h * 0.25), int(w * 0.375),
                            int(h * 0.50)))
            right = im.crop((int(w * 0.80), int(h * 0.64),
                             int(w * 1.0), int(h * 0.82)))
            left = left.resize((left.width * 2, left.height * 2))
            right = right.resize((right.width * 2, right.height * 2))
            lb, rb = io.BytesIO(), io.BytesIO()
            left.save(lb, format="PNG")
            right.save(rb, format="PNG")
            with self._lock:
                self._pending = (lb.getvalue(), rb.getvalue())
            self._ensure_thread()
            self._event.set()
        except Exception:
            pass

    def values(self) -> dict:
        with self._lock:
            return dict(self._values)

    def _ensure_thread(self) -> None:
        if self._thread is not None and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def _loop(self) -> None:
        while not self._stop:
            if not self._event.wait(0.2):
                continue
            self._event.clear()
            with self._lock:
                data = self._pending
                self._pending = None
            if not data:
                continue
            try:
                left_lines = _ocr_lines(data[0])
                right_lines = _ocr_lines(data[1])
                result = _parse_hud(left_lines, right_lines)
            except Exception as exc:
                print(f"[ocr] HUD 识别失败（忽略）：{exc!r}")
                result = {}
            with self._lock:
                self._values = result


class SelectedNameOCR:
    """后台 OCR：submit 最新画面，text() 读最近一次结果。"""

    def __init__(self):
        self._lock = threading.Lock()
        self._pending: bytes | None = None
        self._text = ""
        self._event = threading.Event()
        self._stop = False
        self._thread: threading.Thread | None = None

    def submit(self, image) -> None:
        try:
            from PIL import Image

            h, w = image.shape[:2]
            # 抓屏流是 BGR；转成 RGB 再 OCR。
            if image.ndim == 3 and image.shape[2] >= 3:
                image = np.ascontiguousarray(image[:, :, 2::-1])
            crop = Image.fromarray(image).crop(
                (0, int(h * 0.289), int(w * 0.208), int(h * 0.382)))
            crop = crop.resize((crop.width * 2, crop.height * 2))
            buf = io.BytesIO()
            crop.save(buf, format="PNG")
            with self._lock:
                self._pending = buf.getvalue()
            self._ensure_thread()
            self._event.set()
        except Exception:
            pass

    def text(self) -> str:
        with self._lock:
            return self._text

    def _ensure_thread(self) -> None:
        if self._thread is not None and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def _loop(self) -> None:
        while not self._stop:
            if not self._event.wait(0.2):
                continue
            self._event.clear()
            with self._lock:
                data = self._pending
                self._pending = None
            if not data:
                continue
            try:
                text = _ocr_png(data)
            except Exception as exc:
                print(f"[ocr] 识别失败（忽略）：{exc!r}")
                text = ""
            with self._lock:
                self._text = text
