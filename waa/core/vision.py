"""干员头像识别（本地头像资源 + 模板匹配）。"""

from __future__ import annotations

import glob
import os
import threading

import cv2
import numpy as np

from .resources import (AVATAR_ROOT, OPERATOR_MAPPING_PATH, RESOURCE_LOCK,
                        ROOT, load_operator_mapping)
from ..paths import bundle_root

OPERATOR_AREA_RATIO = (0.0, 0.8, 1.0, 1.0)
AVATAR_STANDARD_SIZE = (120, 120)
AVATAR_CROP_SIZE = (60, 60)
MATCH_THRESHOLD = 0.80
STANDARD_SIZE = (1280, 720)
# 待部署区干员太多时，游戏会把所有头像缩小。固定 60px 模板会匹配
# 不到，所以按多档缩放一起找；实际尺寸在 100%~60% 之间。
AVATAR_SCALES = tuple(round(0.60 + 0.04 * i, 2) for i in range(11))

# MAA 6.21.1 的待部署卡定位方式（resource/tasks/tasks.json）：
#   BattleOpersFlag  : 10x11 小旗标，1280x720 下 ROI [35,588,1245,18]，阈值 0.65
#   BattleOperClickRange: 相对旗标 [-45, 6, 75, 120]
#   BattleOperAvatar   : 相对卡片 [7, 32, 60, 60]
# MAA 本身没有 60%~100% 的头像缩放档位；这里用它找卡，再在卡片
# 头像区域内做尺寸搜索，卡的点击位置不再依赖猜出来的头像中心。
VISION_RESOURCE = os.path.join(bundle_root(), "waa", "resources")
FLAG_TEMPLATE_PATH = os.path.join(VISION_RESOURCE, "BattleOpersFlag.png")
FLAG_ROI = (35, 588, 1245, 18)
FLAG_THRESHOLD = 0.65
OPER_CLICK_MOVE = (-45, 6, 75, 120)
OPER_AVATAR_MOVE = (7, 32, 60, 60)
# 实测 1-7 五张卡时旗标间距约 118.5px（1280x720）；单位更多时
# 间距会变小，用这个比例缩放卡片和头像的搜索尺寸。
STANDARD_FLAG_PITCH = 118.5
CARD_AVATAR_SCALES = (0.70, 0.80, 0.90, 1.00)

_mapping: dict[str, str] | None = None
_avatar_cache: dict[str, list[np.ndarray]] = {}
_avatar_scaled_cache: dict[tuple[str, float], list[np.ndarray]] = {}
_flag_template: np.ndarray | None = None
# 整区识别：待部署卡实际头像约 60px（不随卡片间距一起缩小），
# 统一缩到 40px 做匹配，14~15 张卡一次约 1~2 秒。
CARD_MATCH_SIZE = 40
CARD_MATCH_SCALE = CARD_MATCH_SIZE / 60.0
CARD_MATCH_THRESHOLD = 0.70
# 卡片头像定位最低分；高到接近像素级完全一致时可以直接信任，
# 否则还要核对这张卡在全头像库里的第一名，避免近似头像误认。
CARD_MIN_MATCH = 0.70
CARD_DIRECT_MATCH = 0.995
_card_templates: list[tuple[str, np.ndarray]] | None = None
_card_build_lock = threading.Lock()
_avatar_generation = 0
CARD_CACHE_VERSION = 1
CARD_TEMPLATE_CACHE_PATH = os.path.join(
    ROOT, "out", "card_templates_v1.npz")


def operator_mapping() -> dict[str, str]:
    global _mapping
    with RESOURCE_LOCK:
        if _mapping is None:
            _mapping = load_operator_mapping()
        return _mapping


def _process_avatar(path: str) -> np.ndarray:
    avatar = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
    if avatar is None:
        raise FileNotFoundError(path)
    avatar = cv2.resize(avatar, AVATAR_STANDARD_SIZE)
    y, x = avatar.shape
    cropx, cropy = AVATAR_CROP_SIZE
    startx = x // 2 - cropx // 2
    starty = y // 2 - cropy // 2
    return avatar[starty:starty + cropy, startx:startx + cropx]


def get_avatars(oper: str) -> list[np.ndarray]:
    with RESOURCE_LOCK:
        if oper in _avatar_cache:
            return _avatar_cache[oper]
        mapping = operator_mapping()
        if oper not in mapping:
            raise KeyError(f"干员表里没有: {oper}")
        fragment = mapping[oper]
        paths = glob.glob(os.path.join(AVATAR_ROOT, f"*{fragment}*"))
        if not paths:
            raise FileNotFoundError(f"找不到头像: {oper} -> {fragment}")
        avatars = [_process_avatar(p) for p in paths]
        _avatar_cache[oper] = avatars
        return avatars


def clear_avatar_cache() -> None:
    global _mapping, _card_templates, _avatar_generation
    with RESOURCE_LOCK:
        _avatar_generation += 1
        _mapping = None
        _avatar_cache.clear()
        _avatar_scaled_cache.clear()
        _card_templates = None


def _scaled_avatars(oper: str, scale: float) -> list[np.ndarray]:
    with RESOURCE_LOCK:
        key = (oper, scale)
        cached = _avatar_scaled_cache.get(key)
        if cached is not None:
            return cached
        if scale == 1.0:
            scaled = get_avatars(oper)
        else:
            scaled = [
                cv2.resize(a, None, fx=scale, fy=scale,
                           interpolation=cv2.INTER_AREA)
                for a in get_avatars(oper)
            ]
        _avatar_scaled_cache[key] = scaled
        return scaled


def _get_flag_template() -> np.ndarray | None:
    global _flag_template
    if _flag_template is None:
        _flag_template = cv2.imread(FLAG_TEMPLATE_PATH, cv2.IMREAD_GRAYSCALE)
    return _flag_template


def find_operator_flags(frame_bgr: np.ndarray) -> list[tuple[float, float, float]]:
    """找待部署卡顶部的小旗标，返回 (x, y, score)，坐标是 1280x720。

    这是 MAA 的 BattleOpersFlag 做法；单位多、头像缩小时旗标仍能
    定位每张卡，不依赖头像尺寸。
    """
    tpl = _get_flag_template()
    if tpl is None:
        return []
    frame = cv2.resize(frame_bgr, STANDARD_SIZE, interpolation=cv2.INTER_AREA)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    x0, y0, w, h = FLAG_ROI
    x1 = min(STANDARD_SIZE[0], x0 + w)
    y1 = min(STANDARD_SIZE[1], y0 + h)
    roi = gray[y0:y1, x0:x1]
    if roi.shape[0] < tpl.shape[0] or roi.shape[1] < tpl.shape[1]:
        return []
    result = cv2.matchTemplate(roi, tpl, cv2.TM_CCOEFF_NORMED)
    ys, xs = np.where(result >= FLAG_THRESHOLD)
    candidates = sorted(
        ((float(result[y, x]), int(x) + x0, int(y) + y0)
         for y, x in zip(ys, xs)),
        key=lambda item: item[0], reverse=True)
    flags: list[tuple[float, float, float]] = []
    for score, x, y in candidates:
        # 同一张卡顶部会有相邻的重复峰值，按 x 合并。
        if any(abs(x - fx) < 25 and abs(y - fy) < 15
               for fx, fy, _ in flags):
            continue
        flags.append((float(x), float(y), float(score)))
    flags.sort(key=lambda item: item[0])
    return flags


def operator_cards(frame_bgr: np.ndarray) -> list[dict]:
    """按 MAA 旗标几何返回每张待部署卡的点击点/头像框。

    返回 dict：click=(cx,cy), avatar=(x,y,w,h), scale, flag_score。
    """
    frame = cv2.resize(frame_bgr, STANDARD_SIZE, interpolation=cv2.INTER_AREA)
    flags = find_operator_flags(frame)
    if not flags:
        return []
    xs = [f[0] for f in flags]
    if len(xs) >= 2:
        pitch = float(np.median(np.diff(sorted(xs))))
    else:
        pitch = STANDARD_FLAG_PITCH
    scale = max(0.50, min(1.10, pitch / STANDARD_FLAG_PITCH))
    cards = []
    for fx, fy, score in flags:
        cx0 = fx + OPER_CLICK_MOVE[0] * scale
        cy0 = fy + OPER_CLICK_MOVE[1] * scale
        cw = OPER_CLICK_MOVE[2] * scale
        ch = OPER_CLICK_MOVE[3] * scale
        ax = cx0 + OPER_AVATAR_MOVE[0] * scale
        ay = cy0 + OPER_AVATAR_MOVE[1] * scale
        aw = OPER_AVATAR_MOVE[2] * scale
        ah = OPER_AVATAR_MOVE[3] * scale
        cards.append({
            "click": (cx0 + cw / 2.0, cy0 + ch / 2.0),
            "avatar": (ax, ay, aw, ah),
            "scale": scale,
            "flag_score": score,
        })
    return cards


def rightmost_card_info(frame_bgr: np.ndarray) -> dict | None:
    """返回最右待部署卡（含头像中心）；找不到旗标时返回 None。"""
    cards = operator_cards(frame_bgr)
    if not cards:
        return None
    for card in cards:
        ax, ay, aw, ah = card["avatar"]
        card["avatar_center"] = (float(ax + aw / 2.0),
                                 float(ay + ah / 2.0))
    return max(cards, key=lambda item: item["avatar_center"][0])


def rightmost_card_center(frame_bgr: np.ndarray) -> tuple[float, float] | None:
    """返回最右待部署卡的点击比例；找不到旗标时返回 None。"""
    card = rightmost_card_info(frame_bgr)
    if card is None:
        return None
    x, y = card["click"]
    return x / STANDARD_SIZE[0], y / STANDARD_SIZE[1]


def _match_target_card(gray: np.ndarray, card: dict,
                       oper: str) -> float:
    """在单张卡的卡面区域内找目标干员头像，返回最高分。

    同时兼容两种真实情况：
      - 卡面显示 60px 头像，不随卡片间距缩小（13/14 张卡）；
      - 整张卡等比缩小，头像只有 40px 左右（离线合成/低分辨率）。
    所以既在整张卡面里试 60/40px，也在旗标算出的头像框里试缩放档位。
    """
    cx, cy = card["click"]
    x0 = max(0, int(cx - 42))
    x1 = min(STANDARD_SIZE[0], int(cx + 42))
    y0 = max(0, int(cy - 60))
    y1 = min(STANDARD_SIZE[1], int(cy + 60))
    raw = gray[y0:y1, x0:x1]
    if raw.size == 0:
        return 0.0
    ax, ay, aw, ah = card["avatar"]
    ax0 = max(0, int(round(ax)))
    ay0 = max(0, int(round(ay)))
    ax1 = min(STANDARD_SIZE[0], ax0 + max(8, int(round(aw))))
    ay1 = min(STANDARD_SIZE[1], ay0 + max(8, int(round(ah))))
    avatar_box = gray[ay0:ay1, ax0:ax1]
    sizes = {60, CARD_MATCH_SIZE}
    if aw >= 10:
        sizes.add(int(round(aw)))
    best = 0.0
    avatars = get_avatars(oper)
    for avatar in avatars:
        for size in sizes:
            if size <= 0 or size > raw.shape[0] or size > raw.shape[1]:
                continue
            tpl = cv2.resize(avatar, (size, size),
                             interpolation=cv2.INTER_AREA)
            score = float(cv2.matchTemplate(
                raw, tpl, cv2.TM_CCOEFF_NORMED).max())
            if score > best:
                best = score
        if avatar_box.size:
            for f in CARD_AVATAR_SCALES:
                tw = max(8, int(round(aw * f)))
                th = max(8, int(round(ah * f)))
                if tw > avatar_box.shape[1] or th > avatar_box.shape[0]:
                    continue
                tpl = cv2.resize(avatar, (tw, th),
                                 interpolation=cv2.INTER_AREA)
                score = float(cv2.matchTemplate(
                    avatar_box, tpl, cv2.TM_CCOEFF_NORMED).max())
                if score > best:
                    best = score
    return best


def _best_card_identity(gray: np.ndarray, card: dict
                        ) -> tuple[str | None, float]:
    """对一张卡跑完整头像库，返回分数最高的名字和分数。"""
    crop = _card_crop(gray, card)
    if crop is None or crop.size == 0:
        return None, 0.0
    best_name = None
    best_score = 0.0
    for name, template in _get_card_templates():
        if (template.shape[0] > crop.shape[0]
                or template.shape[1] > crop.shape[1]):
            continue
        score = float(cv2.matchTemplate(
            crop, template, cv2.TM_CCOEFF_NORMED).max())
        if score > best_score:
            best_score = score
            best_name = name
    return best_name, best_score


def _same_operator(a: str | None, b: str | None) -> bool:
    if not a or not b:
        return False
    if a == b:
        return True
    mapping = operator_mapping()
    return bool(mapping.get(a) and mapping.get(a) == mapping.get(b))


def locate_avatar(frame_bgr: np.ndarray, oper: str) -> tuple[float, float, float]:
    """返回 (x_ratio, y_ratio, score)，坐标是 1280x720 标准画面下卡片中心。

    优先用 MAA 的旗标找到每张待部署卡，再只在卡片头像框内做尺寸
    匹配；这样单位多、头像缩小时，点击点仍由卡片位置决定，不需要
    在整个待部署区盲搜头像。找不到旗标时才退回全区域匹配。
    """
    frame = cv2.resize(frame_bgr, STANDARD_SIZE, interpolation=cv2.INTER_AREA)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

    cards = operator_cards(frame)
    if cards:
        # 卡片实际头像约 60px，不能按卡片框缩放后的 46px 去匹配。
        # 先在每张卡的卡面区域里找目标，再对分数不高的卡做一次
        # “全头像库第一名”核对，避免把相似头像认成目标。
        scored = sorted(
            ((_match_target_card(gray, card, oper), card)
             for card in cards),
            key=lambda item: item[0], reverse=True)
        best_score = scored[0][0] if scored else 0.0
        for score, card in scored[:2]:
            if score < CARD_MIN_MATCH:
                break
            # 接近像素级完全一致才直接信任；否则核对“全头像库第一名”，
            # 不同形态/近似头像（如电弧与 Raidian）会互相误认。
            if score < CARD_DIRECT_MATCH:
                top_name, top_score = _best_card_identity(gray, card)
                print(f"  avatar verify {oper}: card_score={score:.3f} "
                      f"top={top_name} top_score={top_score:.3f}")
                if not (_same_operator(top_name, oper)
                        and top_score >= CARD_MIN_MATCH):
                    continue
            cx, cy = card["click"]
            print(f"  avatar card-match cards={len(cards)} "
                  f"scale={card['scale']:.2f} "
                  f"score={score:.3f} flag={card['flag_score']:.3f}")
            return (float(cx / STANDARD_SIZE[0]),
                    float(cy / STANDARD_SIZE[1]), float(score))
        raise RuntimeError(
            f"待部署区没找到 {oper}，最高卡片匹配 {best_score:.3f}")

    # 没有旗标时的兜底：全待部署区多尺寸匹配。
    l, t, r, b = OPERATOR_AREA_RATIO
    area = gray[int(t * STANDARD_SIZE[1]):int(b * STANDARD_SIZE[1]),
                int(l * STANDARD_SIZE[0]):int(r * STANDARD_SIZE[0])]

    best_score = 0.0
    best_pos = None
    best_avatar = None
    best_scale = 1.0
    for scale in AVATAR_SCALES:
        for avatar in _scaled_avatars(oper, scale):
            if (avatar.shape[0] > area.shape[0]
                    or avatar.shape[1] > area.shape[1]):
                continue
            result = cv2.matchTemplate(area, avatar, cv2.TM_CCOEFF_NORMED)
            _, score, _, pos = cv2.minMaxLoc(result)
            if score > best_score:
                best_score = score
                best_pos = pos
                best_avatar = avatar
                best_scale = scale

    if best_pos is None or best_score < MATCH_THRESHOLD:
        raise RuntimeError(
            f"待部署区没找到 {oper}，最高匹配 {best_score:.3f}")

    x = (best_pos[0] + best_avatar.shape[1] / 2) / STANDARD_SIZE[0]
    y = t + (best_pos[1] + best_avatar.shape[0] / 2) / STANDARD_SIZE[1]
    # 头像全区域匹配成功、但卡片内匹配失败时，把落点吸附到最近的
    # 旗标卡片中心。这样拖拽一定从真实卡片上开始，不会点进卡片
    # 之间的空隙导致“仍在子弹时间”被误判为部署成功。
    if cards:
        px, py = x * STANDARD_SIZE[0], y * STANDARD_SIZE[1]
        nearest = min(
            cards,
            key=lambda c: ((c["click"][0] - px) ** 2
                           + (c["click"][1] - py) ** 2))
        if (abs(nearest["click"][0] - px) <= 0.07 * STANDARD_SIZE[0]
                and abs(nearest["click"][1] - py) <= 0.15 * STANDARD_SIZE[1]):
            x = nearest["click"][0] / STANDARD_SIZE[0]
            y = nearest["click"][1] / STANDARD_SIZE[1]
            print(f"  avatar snap to card "
                  f"cards={len(cards)} scale={nearest['scale']:.2f}")
    print(f"  avatar fallback cards={len(cards)} "
          f"scale={best_scale:.2f} score={best_score:.3f}")
    return float(x), float(y), float(best_score)


def _resource_mtime() -> tuple[float, float]:
    """缓存失效依据：角色映射表 + 头像目录的修改时间。"""
    try:
        mapping_mtime = os.path.getmtime(OPERATOR_MAPPING_PATH)
    except OSError:
        mapping_mtime = 0.0
    try:
        avatar_mtime = os.path.getmtime(AVATAR_ROOT)
    except OSError:
        avatar_mtime = 0.0
    return float(mapping_mtime), float(avatar_mtime)


def _load_card_template_cache(log=None) -> list[tuple[str, np.ndarray]] | None:
    """读取磁盘模板缓存；版本或资源变化则视为失效。"""
    if not os.path.exists(CARD_TEMPLATE_CACHE_PATH):
        return None
    try:
        data = np.load(CARD_TEMPLATE_CACHE_PATH, allow_pickle=True)
        version = int(data["version"])
        mapping_mtime = float(data["mapping_mtime"])
        avatar_mtime = float(data["avatar_mtime"])
        current = _resource_mtime()
        if (version != CARD_CACHE_VERSION
                or (mapping_mtime, avatar_mtime) != current):
            return None
        templates = np.asarray(data["templates"])
        names = data["names"]
        if templates.ndim != 3 or templates.shape[1:] != (
                CARD_MATCH_SIZE, CARD_MATCH_SIZE):
            return None
        if len(names) != templates.shape[0]:
            return None
        result = [(str(names[i]), np.ascontiguousarray(templates[i]))
                  for i in range(templates.shape[0])]
        if log:
            log(f"头像库缓存已加载：{len(result)} 个模板")
        return result
    except Exception as exc:
        if log:
            log(f"头像库缓存不可用（将重新加载）：{exc!r}")
        return None


def _save_card_template_cache(templates: list[tuple[str, np.ndarray]],
                              log=None) -> None:
    if not templates:
        return
    try:
        os.makedirs(os.path.dirname(CARD_TEMPLATE_CACHE_PATH), exist_ok=True)
        names = np.array([name for name, _ in templates], dtype=object)
        arrays = np.stack([tpl for _, tpl in templates]).astype(np.uint8)
        mapping_mtime, avatar_mtime = _resource_mtime()
        tmp = CARD_TEMPLATE_CACHE_PATH + ".tmp.npz"
        np.savez_compressed(
            tmp, version=CARD_CACHE_VERSION,
            mapping_mtime=mapping_mtime, avatar_mtime=avatar_mtime,
            names=names, templates=arrays)
        os.replace(tmp, CARD_TEMPLATE_CACHE_PATH)
        if log:
            log(f"头像库缓存已保存：{CARD_TEMPLATE_CACHE_PATH}")
    except Exception as exc:
        if log:
            log(f"头像库缓存保存失败（不影响运行）：{exc!r}")


def _get_card_templates(log=None) -> list[tuple[str, np.ndarray]]:
    """所有本地干员头像的 40×40 灰度模板，用于整区卡片识别。

    同一个头像 fragment 只保留一个名字；首次加载约 10 秒，之后复用。
    加载过程不长时间占着 RESOURCE_LOCK，避免阻塞 load/地图/头像更新。
    """
    global _card_templates
    with RESOURCE_LOCK:
        if _card_templates is not None:
            return _card_templates
    with _card_build_lock:
        with RESOURCE_LOCK:
            if _card_templates is not None:
                return _card_templates
            mapping = dict(operator_mapping())
            generation = _avatar_generation
        cached = _load_card_template_cache(log=log)
        if cached:
            with RESOURCE_LOCK:
                _card_templates = cached
            return cached
        if log:
            log("正在加载待部署头像库（首次约 10 秒）…")
        else:
            print("  [vision] 首次加载待部署头像库，约 10 秒…")
        seen_fragments = set()
        templates = []
        for name, fragment in mapping.items():
            if fragment in seen_fragments:
                continue
            seen_fragments.add(fragment)
            try:
                avatars = get_avatars(name)
            except Exception:
                continue
            for avatar in avatars:
                templates.append((name, cv2.resize(
                    avatar, (CARD_MATCH_SIZE, CARD_MATCH_SIZE),
                    interpolation=cv2.INTER_AREA)))
        with RESOURCE_LOCK:
            # 加载过程中如果资源被更新，缓存可能会是旧数据；这次仍返回，
            # 但清掉缓存，下一次调用会重建。
            if generation == _avatar_generation:
                _card_templates = templates
            else:
                _card_templates = None
        if generation == _avatar_generation:
            _save_card_template_cache(templates, log=log)
        if log:
            log(f"头像库已加载：{len(templates)} 个模板")
        return templates


def preload_card_templates(log=None, force: bool = False) -> bool:
    """启动时只读磁盘缓存，不重新扫描/构建头像库。

    缓存不存在或资源已更新时直接返回，等第一次 cards/部署时再构建，
    避免每次打开控制台都花约 10 秒。

    `force=True`：缓存缺失就**现场构建**。run 之前必须这么做——首次识别
    发生在选人/部署的帧窗口里（几百毫秒），而构建要 ~10 秒，边跑边加载
    会直接毁掉那一行。
    """
    global _card_templates
    if force:
        with RESOURCE_LOCK:
            if _card_templates is not None:
                return True
        try:
            _get_card_templates(log=log)
        except Exception as exc:
            if log:
                log(f"头像库加载失败：{exc}")
            return False
        with RESOURCE_LOCK:
            return _card_templates is not None
    with RESOURCE_LOCK:
        if _card_templates is not None:
            return True
    with _card_build_lock:
        with RESOURCE_LOCK:
            if _card_templates is not None:
                return True
        cached = _load_card_template_cache(log=log)
        if cached:
            with RESOURCE_LOCK:
                _card_templates = cached
            return True
        if log:
            log("头像缓存还没建立：首次识别会自动扫描头像库（约 10 秒），"
                "之后启动直接读缓存")
        return False


def _card_crop(gray: np.ndarray, card: dict) -> np.ndarray | None:
    """取卡片周围一块区域；太靠屏幕边的卡自动裁到边界内。"""
    cx, cy = card["click"]
    x0 = max(0, int(cx - 42))
    x1 = min(STANDARD_SIZE[0], int(cx + 42))
    y0 = max(0, int(cy - 60))
    y1 = min(STANDARD_SIZE[1], int(cy + 60))
    crop = gray[y0:y1, x0:x1]
    if crop.size == 0:
        return None
    return cv2.resize(crop, None, fx=CARD_MATCH_SCALE, fy=CARD_MATCH_SCALE,
                      interpolation=cv2.INTER_AREA)


def identify_cards(frame_bgr: np.ndarray,
                   threshold: float = CARD_MATCH_THRESHOLD,
                   log=None) -> list[dict]:
    """识别当前待部署区的每一张卡，返回名称/分数/位置。

    每张卡都带：
      oper            识别出的单位名；低于 threshold 时为 None
      candidate       分数最高的候选（即使低于阈值）
      identify_score  最高匹配分
      avatar_center   头像中心（1280×720 标准坐标）
      click           卡片点击点（1280×720 标准坐标）
    卡的数量和位置仍由 MAA 旗标检测给出；头像匹配只负责“是谁”。
    """
    frame = cv2.resize(frame_bgr, STANDARD_SIZE,
                       interpolation=cv2.INTER_AREA)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    cards = operator_cards(frame)
    templates = _get_card_templates(log)
    for card in cards:
        crop = _card_crop(gray, card)
        ax, ay, aw, ah = card["avatar"]
        card["avatar_center"] = (float(ax + aw / 2.0),
                                 float(ay + ah / 2.0))
        card["oper"] = None
        card["candidate"] = None
        card["identify_score"] = 0.0
        if crop is None or crop.size == 0:
            continue
        best_score = 0.0
        best_name = None
        for name, template in templates:
            if (template.shape[0] > crop.shape[0]
                    or template.shape[1] > crop.shape[1]):
                continue
            score = float(cv2.matchTemplate(
                crop, template, cv2.TM_CCOEFF_NORMED).max())
            if score > best_score:
                best_score = score
                best_name = name
        card["candidate"] = best_name
        card["identify_score"] = best_score
        if best_name is not None and best_score >= threshold:
            card["oper"] = best_name
    return cards
