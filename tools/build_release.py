"""一键打包：PyInstaller -> release/WAA-v<版本>/，再压同名 zip。

为什么逻辑放在 Python 里：.bat 在中文 Windows 上按 OEM 代码页解析，文件里任何
UTF-8 中文都会让 cmd 拆错行（表现为「'xxx' 不是内部或外部命令」）。所以
build_exe.bat 只是一层纯 ASCII 的壳，真正的构建在这里。

为什么不做「先挪走旧目录」：移动 1000+ 文件的目录会被批量删除保护拦下
（等于一次大量文件消失）。这里改成构建到新的 build/dist_<时间戳>，再逐文件
覆盖到 release/WAA-v<版本>。同一版本重复构建时旧文件不会被清掉——需要干净重建时，
手动把那个版本目录删掉再跑。

用法：
    python tools/build_release.py
    python tools/build_release.py --with-resources
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PY = ROOT / ".venv" / "Scripts" / "python.exe"
DIST = ROOT / "release"
BUILD = ROOT / "build"
PKG = ROOT / "packaging"
README = "使用说明.txt"
TEMPLATE = ROOT / "axes" / "WAA_标准模板.xlsx"


def app_version() -> str:
    """从 waa/__init__.py 读版本号，产物目录跟着它走。"""
    import re
    text = (ROOT / "waa" / "__init__.py").read_text(encoding="utf-8")
    m = re.search(r'__version__\s*=\s*"([^"]+)"', text)
    if not m:
        raise SystemExit("读不到 waa/__init__.py 里的 __version__")
    return m.group(1)


def log(msg: str) -> None:
    print(msg, flush=True)


def run(cmd: list[object]) -> None:
    log("  > " + " ".join(str(c) for c in cmd))
    subprocess.run([str(c) for c in cmd], cwd=str(ROOT), check=True)


def build(stamp: str) -> Path:
    """构建到 build/dist_<stamp>，不碰已有产物。"""
    out = BUILD / f"dist_{stamp}"
    log("[1/4] 构建 WAA_frame.exe（悬浮帧窗）")
    run([PY, "-m", "PyInstaller", "--noconfirm", "--onefile", "--windowed",
         "--name", "WAA_frame", "--distpath", out,
         "--workpath", "build/pyinstaller_frame", "tools/frame_window.py"])

    log("[2/4] 构建 WAA.exe（主控制台）")
    run([PY, "-m", "PyInstaller", "--noconfirm", "--onedir", "--console",
         "--name", "WAA", "--distpath", out,
         "--workpath", "build/pyinstaller", "--paths", ".",
         "--add-data", "waa/resources;waa/resources",
         "--add-data", "waa/templates;waa/templates",
         "tools/axis_console.py"])
    return out


def ship_card_cache(app: Path) -> None:
    """把头像模板缓存一起打进包，并让它「看起来有效」。

    缓存有效性只比对两个 mtime：out/gamedata/operator_mapping.json 和
    out/ArknightsGameResource/avatar 目录。复制过来后 mtime 变了，缓存会被判失效，
    所以这里按缓存里存的值把两个 mtime 还原回去。判错了也无害——程序会自己重建。
    """
    npz = ROOT / "out" / "card_templates_v1.npz"
    target = app / "out" / npz.name
    if not npz.exists():
        return
    try:
        import numpy as np
        data = np.load(npz, allow_pickle=True)
        mapping_mtime = float(data["mapping_mtime"])
        avatar_mtime = float(data["avatar_mtime"])
        target.parent.mkdir(parents=True, exist_ok=True)
        if not target.exists():
            shutil.copy2(npz, target)
        os.utime(app / "out" / "gamedata" / "operator_mapping.json",
                 (mapping_mtime, mapping_mtime))
        os.utime(app / "out" / "ArknightsGameResource" / "avatar",
                 (avatar_mtime, avatar_mtime))
        log(f"  头像缓存已随包（{len(data['names'])} 个模板）")
    except Exception as exc:
        log(f"  [跳过] 头像缓存没打进去：{exc!r}")


def assemble(out: Path, with_resources: bool, app: Path) -> None:
    log(f"[3/4] 组装到 release/{app.name}")
    src_app = out / "WAA"
    if not (src_app / "WAA.exe").exists():
        raise SystemExit(f"构建产物缺失：{src_app / 'WAA.exe'}")

    shutil.copytree(src_app, app, dirs_exist_ok=True)
    shutil.copy2(out / "WAA_frame.exe", app / "WAA_frame.exe")
    shutil.copy2(PKG / README, app / README)

    axes = app / "axes"
    axes.mkdir(exist_ok=True)
    if TEMPLATE.exists():
        # 模板只是附赠品：用户可能正在 Excel 里开着它，拷不过去也不算构建失败。
        try:
            shutil.copy2(TEMPLATE, axes / TEMPLATE.name)
        except OSError as exc:
            log(f"  [跳过] 模板没拷进去（文件被占用）：{exc}")

    if with_resources:
        log("[extra] 打包游戏资源（约 330MB）")
        for name in ("ArknightsGameResource", "gamedata"):
            src = ROOT / "out" / name
            if src.exists():
                shutil.copytree(src, app / "out" / name, dirs_exist_ok=True)
        npz = ROOT / "out" / "card_templates_v1.npz"
        if npz.exists():
            shutil.copy2(npz, app / "out" / npz.name)

    ship_card_cache(app)


def make_zip(app: Path) -> Path:
    """把成品目录压成同名 zip，放在 release/ 下，方便直接传 Release。"""
    out = app.parent / f"{app.name}.zip"
    if out.exists():
        out.unlink()
    n = 0
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for base, dirs, files in os.walk(app):
            dirs.sort()
            for name in sorted(files):
                full = os.path.join(base, name)
                rel = os.path.relpath(full, app.parent)
                z.write(full, Path(rel).as_posix())
                n += 1
    log(f"[4/4] 打包 {out.name}（{n} 个文件）")
    return out


def selftest(app: Path) -> int:
    log("")
    log(f"自检：release/{app.name}/WAA.exe --selftest")
    try:
        return subprocess.run([str(app / "WAA.exe"), "--selftest"],
                              cwd=str(app)).returncode
    except Exception as exc:
        log(f"自检没跑起来：{exc!r}")
        return 1


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--with-resources", action="store_true")
    args = ap.parse_args()

    if not PY.exists():
        log("[错误] 找不到 .venv。先执行：")
        log("        py -3.13 -m venv .venv")
        log("        .venv\\Scripts\\python.exe -m pip install -r requirements.txt")
        return 1

    version = app_version()
    app = DIST / f"WAA-v{version}"
    stamp = time.strftime("%Y%m%d_%H%M%S")
    out = build(stamp)
    assemble(out, args.with_resources, app)

    size_mb = sum(f.stat().st_size for f in app.rglob("*") if f.is_file()) // 1048576
    log("")
    log(f"完成：release/{app.name}/  （{size_mb} MB，构建目录 build/dist_{stamp}）")
    log(f"  启动：双击 release/{app.name}/WAA.exe（会弹 UAC 提权）")

    zip_path = make_zip(app)
    log(f"  发布包：release/{zip_path.name}（{zip_path.stat().st_size // 1048576} MB）")

    if selftest(app) != 0:
        log("")
        log("[注意] 自检未全部通过，见上面 FAIL 行。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
