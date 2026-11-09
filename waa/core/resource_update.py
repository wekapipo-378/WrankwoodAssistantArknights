"""从 yuanyan3060/ArknightsGameResource 更新本地地图/头像/角色表。

运行时只读 out/ArknightsGameResource 和 out/gamedata；这里负责把这两个
目录从上游仓库同步到最新。更新过程中持有 RESOURCE_LOCK，避免和 read /
部署同时读写资源。
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import threading
import time

from .resources import (
    GAMEDATA_ROOT, GIT_REMOTE, GIT_REPO_PATH, OPERATOR_MAPPING_PATH,
    RESOURCE_LOCK, UPDATE_STATE_PATH, atomic_write_bytes, atomic_write_json,
    build_operator_mapping,
)

AUTO_INTERVAL = 6 * 3600
GIT_TIMEOUT = 180


class ResourceUpdater:
    def __init__(self, repo: str = GIT_REPO_PATH):
        self.repo = repo
        self._thread: threading.Thread | None = None

    # ---------- 状态 ----------
    @staticmethod
    def _load_state() -> dict:
        try:
            with open(UPDATE_STATE_PATH, encoding="utf-8") as fh:
                data = json.load(fh)
            return data if isinstance(data, dict) else {}
        except (OSError, json.JSONDecodeError):
            return {}

    @staticmethod
    def _save_state(**values) -> None:
        state = ResourceUpdater._load_state()
        state.update(values)
        try:
            atomic_write_json(UPDATE_STATE_PATH, state)
        except OSError:
            pass

    # ---------- git ----------
    @staticmethod
    def _find_git() -> str | None:
        """找 git.exe：PATH 优先，其次常见安装位置和便携版。

        控制台是管理员启动的，PATH 可能不含用户级安装的 Git；本机也可能
        只有 WorkBuddy 自带的便携版，所以这里多找几处再报错。
        """
        exe = shutil.which("git")
        if exe:
            return exe
        candidates: list[str] = []
        for env in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
            base = os.environ.get(env)
            if not base:
                continue
            candidates += [
                os.path.join(base, "Git", "cmd", "git.exe"),
                os.path.join(base, "Git", "bin", "git.exe"),
                os.path.join(base, "Programs", "Git", "cmd", "git.exe"),
            ]
        # WorkBuddy 自带便携版
        candidates += [
            r"C:\Program Files\WorkBuddy\resources\vendor\PortableGit\mingw64\bin\git.exe",
            r"C:\Program Files\WorkBuddy\resources\vendor\PortableGit\cmd\git.exe",
        ]
        home = os.path.expanduser("~")
        vroot = os.path.join(home, ".workbuddy", "binaries",
                             "PortableGit", "versions")
        try:
            versions = sorted(os.listdir(vroot), reverse=True)
        except OSError:
            versions = []
        for name in versions:
            candidates.append(os.path.join(vroot, name, "bin", "git.exe"))
            candidates.append(os.path.join(vroot, name, "cmd", "git.exe"))
        for path in candidates:
            if path and os.path.isfile(path):
                return path
        return None

    def _git(self, *args: str, timeout: int = GIT_TIMEOUT,
             text: bool = True):
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        exe = getattr(self, "git_exe", None) or self._find_git() or "git"
        cmd = [exe, "-C", self.repo, *args]
        return subprocess.run(
            cmd, capture_output=True, text=text, timeout=timeout,
            creationflags=flags)

    def _origin_ok(self) -> bool:
        try:
            proc = self._git("remote", "get-url", "origin", timeout=20)
        except Exception:
            return False
        url = (proc.stdout or "").strip()
        if proc.returncode != 0:
            return False
        return "yuanyan3060/ArknightsGameResource" in url

    def _local_head(self) -> str:
        proc = self._git("rev-parse", "HEAD", timeout=20)
        if proc.returncode != 0:
            raise RuntimeError(f"读取本地版本失败: {(proc.stderr or '').strip()}")
        return (proc.stdout or "").strip()

    def _remote_head(self) -> str:
        proc = self._git("ls-remote", "origin", "refs/heads/main", timeout=30)
        if proc.returncode != 0:
            raise RuntimeError(f"连接 GitHub 失败: {(proc.stderr or '').strip()}")
        line = (proc.stdout or "").strip().splitlines()
        if not line:
            raise RuntimeError("GitHub 没有返回 main 分支")
        return line[0].split()[0]

    def _is_dirty(self) -> bool:
        proc = self._git("status", "--porcelain", timeout=20)
        return proc.returncode == 0 and bool((proc.stdout or "").strip())

    # ---------- gamedata / mapping ----------
    def _extract_gamedata(self) -> None:
        for name in ("character_table.json", "token_table.json"):
            proc = self._git(
                "show", f"HEAD:gamedata/excel/{name}",
                timeout=GIT_TIMEOUT, text=False)
            if proc.returncode != 0:
                err = (proc.stderr or b"").decode("utf-8", "replace").strip()
                raise RuntimeError(f"提取 {name} 失败: {err}")
            atomic_write_bytes(os.path.join(GAMEDATA_ROOT, name),
                               proc.stdout or b"")

    def _refresh_mapping(self) -> int:
        existing = {}
        if os.path.exists(OPERATOR_MAPPING_PATH):
            try:
                with open(OPERATOR_MAPPING_PATH, encoding="utf-8") as fh:
                    existing = json.load(fh)
            except (OSError, json.JSONDecodeError):
                existing = {}
        mapping = build_operator_mapping(
            existing=existing if isinstance(existing, dict) else {},
            write=True)
        return len(mapping)

    @staticmethod
    def _clear_caches() -> None:
        try:
            from .axis import clear_map_cache
            clear_map_cache()
        except Exception:
            pass
        try:
            from .vision import clear_avatar_cache
            clear_avatar_cache()
        except Exception:
            pass

    # ---------- 主流程 ----------
    def update(self, force: bool = False, quiet: bool = False) -> dict:
        """返回 {status, old, new, message, count}。"""
        with RESOURCE_LOCK:
            state = self._load_state()
            now = time.time()
            if not force:
                last = float(state.get("last_check") or 0)
                if last and now - last < AUTO_INTERVAL:
                    return {"status": "skipped", "message": "距离上次检查未满 6 小时"}

            try:
                if self._find_git() is None:
                    raise RuntimeError(
                        "找不到 git：已查 PATH、Program Files\\Git、"
                        "WorkBuddy 自带 PortableGit。请安装 Git for Windows，"
                        "或把 git.exe 所在目录加到 PATH")
                self.git_exe = self._find_git()
                if not os.path.isdir(os.path.join(self.repo, ".git")):
                    raise RuntimeError(f"找不到本地数据仓库: {self.repo}")
                if not self._origin_ok():
                    raise RuntimeError(
                        "本地数据仓库的 origin 不是 "
                        "yuanyan3060/ArknightsGameResource")
                if self._is_dirty():
                    raise RuntimeError(
                        "本地数据目录有未提交修改，已跳过自动更新")

                old = self._local_head()
                remote = self._remote_head()
                if old == remote:
                    # 即使代码版本没变，也补齐可能缺失的本地派生文件。
                    self._extract_gamedata()
                    count = self._refresh_mapping()
                    self._clear_caches()
                    self._save_state(last_check=now, last_head=old,
                                     last_error="")
                    msg = f"已是最新：{old[:12]}"
                    if not quiet:
                        print(f"[update] {msg}")
                    return {"status": "up-to-date", "old": old, "new": old,
                            "count": count, "message": msg}

                proc = self._git("pull", "--ff-only", timeout=GIT_TIMEOUT)
                if proc.returncode != 0:
                    err = (proc.stderr or proc.stdout or "").strip()
                    raise RuntimeError(f"git pull 失败: {err}")

                new = self._local_head()
                self._extract_gamedata()
                count = self._refresh_mapping()
                self._clear_caches()
                self._save_state(last_check=now, last_head=new, last_error="")
                msg = f"已更新：{old[:12]} -> {new[:12]}，角色映射 {count} 条"
                if not quiet:
                    print(f"[update] {msg}")
                return {"status": "updated", "old": old, "new": new,
                        "count": count, "message": msg}
            except Exception as exc:
                self._save_state(last_check=now, last_error=str(exc))
                msg = f"更新失败：{exc}"
                if not quiet:
                    print(f"[update] {msg}")
                return {"status": "error", "message": msg}

    def start_background(self, report=None) -> None:
        """启动后台检查；不阻塞控制台。"""
        if self._thread is not None and self._thread.is_alive():
            return
        self._thread = threading.Thread(
            target=self._background, args=(report,), daemon=True)
        self._thread.start()

    def _background(self, report) -> None:
        result = self.update(force=False, quiet=True)
        if report is None:
            return
        status = result.get("status")
        if status == "updated":
            report(result.get("message", "数据已更新"))
        elif status == "error":
            report(result.get("message", "数据更新失败"))
