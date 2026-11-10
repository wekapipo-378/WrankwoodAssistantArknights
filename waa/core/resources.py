"""本地游戏资源路径与名称映射。

运行时不再读取 prts-plus / MAA 等外部程序目录：
  - 地图：out/ArknightsGameResource/levels.json
  - 头像：out/ArknightsGameResource/avatar
  - 角色/召唤物表：out/gamedata/*.json

RESOURCE_LOCK 在更新资源时加写锁，读地图/头像/映射时加读锁，
避免后台自动更新和 read/部署同时改文件。
"""

from __future__ import annotations

import json
import os
import threading

from ..paths import app_root

ROOT = app_root()

GAME_RESOURCE_ROOT = os.path.join(ROOT, "out", "ArknightsGameResource")
AVATAR_ROOT = os.path.join(GAME_RESOURCE_ROOT, "avatar")
LEVELS_PATH = os.path.join(GAME_RESOURCE_ROOT, "levels.json")
GAMEDATA_ROOT = os.path.join(ROOT, "out", "gamedata")
OPERATOR_MAPPING_PATH = os.path.join(GAMEDATA_ROOT, "operator_mapping.json")

GIT_REPO_PATH = GAME_RESOURCE_ROOT
GIT_REMOTE = "https://github.com/yuanyan3060/ArknightsGameResource.git"
UPDATE_STATE_PATH = os.path.join(ROOT, "out", "resource_update_state.json")

RESOURCE_LOCK = threading.RLock()


def atomic_write_json(path: str, data) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def atomic_write_bytes(path: str, data: bytes) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "wb") as fh:
        fh.write(data)
    os.replace(tmp, path)


def clean_name(name: str) -> str:
    for ch in ("“", "”", "「", "」", "　"):
        name = name.replace(ch, "")
    return name.strip()


def _write_mapping_if_changed(mapping: dict) -> None:
    """内容相同就不写，避免无谓刷新 mtime 导致头像库缓存失效。"""
    if os.path.exists(OPERATOR_MAPPING_PATH):
        try:
            with open(OPERATOR_MAPPING_PATH, encoding="utf-8") as fh:
                old = json.load(fh)
            if old == mapping:
                return
        except (OSError, json.JSONDecodeError):
            pass
    atomic_write_json(OPERATOR_MAPPING_PATH, mapping)


def build_operator_mapping(existing: dict | None = None,
                           write: bool = True) -> dict:
    """从本地 character_table / token_table 重建“名称 -> ID”映射。"""
    with RESOURCE_LOCK:
        mapping = dict(existing or {})
        if not mapping and os.path.exists(OPERATOR_MAPPING_PATH):
            try:
                with open(OPERATOR_MAPPING_PATH, encoding="utf-8") as fh:
                    mapping = json.load(fh)
            except (OSError, json.JSONDecodeError):
                mapping = {}

        for table_name in ("character_table.json", "token_table.json"):
            path = os.path.join(GAMEDATA_ROOT, table_name)
            if not os.path.exists(path):
                continue
            try:
                with open(path, encoding="utf-8") as fh:
                    table = json.load(fh)
            except (OSError, json.JSONDecodeError):
                continue
            if not isinstance(table, dict):
                continue
            for item_id, item in table.items():
                if not isinstance(item, dict):
                    continue
                name = str(item.get("name") or "").strip()
                if not name:
                    continue
                mapping.setdefault(name, item_id)
                cleaned = clean_name(name)
                if cleaned and cleaned != name:
                    mapping.setdefault(cleaned, item_id)

        if write:
            # 只有内容真的变了才写：atomic_write_json 无条件 os.replace 会刷新
            # mtime，即使内容一字未变也把 operator_mapping.json 的 mtime 顶掉，
            # 导致依赖 mtime 对比的头像库缓存（card_templates_v1.npz）被误判
            # 失效，每次启动都要花 ~10 秒重建（2026-11-09 13:03 就是 mapping
            # 被重写、缓存失效）。
            _write_mapping_if_changed(mapping)
        return mapping


def load_operator_mapping() -> dict:
    with RESOURCE_LOCK:
        if os.path.exists(OPERATOR_MAPPING_PATH):
            try:
                with open(OPERATOR_MAPPING_PATH, encoding="utf-8") as fh:
                    data = json.load(fh)
                if isinstance(data, dict):
                    return data
            except (OSError, json.JSONDecodeError):
                pass
        return build_operator_mapping(write=True)
