#!/usr/bin/env python
"""把「到梦空间自动化脚本」打包成一个 Windows exe。

用法（在本文件所在目录执行）::

    python build_exe.py            # 打 onefile 单 exe
    python build_exe.py --onedir   # 打 目录版（启动更快）

打包前会从 tool_assets/platform-tools.zip（或直接联网下载）准备 adb，
再把精简后的 platform-tools 放进包里，最终产物在 dist/ 目录下。
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PKG_DIR = ROOT / "daomeng"
PT_DST = PKG_DIR / "platform-tools"
PT_ZIP = ROOT / "tool_assets" / "platform-tools.zip"
PT_SRC = ROOT / "tool_assets" / "platform-tools"
DOWNLOAD_URL = "https://dl.google.com/android/repository/platform-tools-latest-windows.zip"

# adb 运行所需的最小文件集合（-15MB → -8MB）
NEEDED_FILES = ["adb.exe", "AdbWinApi.dll", "AdbWinUsbApi.dll", "libwinpthread-1.dll", "NOTICE.txt"]


def prepare_platform_tools() -> None:
    """确保 daomeng/platform-tools 下有可用的 adb。"""
    if (PT_DST / "adb.exe").exists():
        print(f"[skip] platform-tools 已就绪: {PT_DST}")
        return

    if not (PT_SRC / "adb.exe").exists():
        PT_ZIP.parent.mkdir(parents=True, exist_ok=True)
        if not PT_ZIP.exists():
            print(f"[down] 下载 platform-tools: {DOWNLOAD_URL}")
            urllib.request.urlretrieve(DOWNLOAD_URL, PT_ZIP)
        print("[unzip] 解压 platform-tools")
        with zipfile.ZipFile(PT_ZIP) as zf:
            zf.extractall(PT_ZIP.parent)

    PT_DST.mkdir(parents=True, exist_ok=True)
    for name in NEEDED_FILES:
        src = PT_SRC / name
        if src.exists():
            shutil.copy2(src, PT_DST / name)
    print(f"[ok] 已准备 adb: {PT_DST / 'adb.exe'}")


VARIANTS = {
    "general": ("DM.py", "到梦空间学分助手"),
    "surname": ("DM_surname.py", "按姓氏批量发分版"),
    "surname-one-point": ("DM_surname.py", "按姓氏批量发分版-只发1分"),
}


def build(onedir: bool = False, variant: str = "general") -> Path:
    prepare_platform_tools()
    entry_name, app_name = VARIANTS[variant]

    sep = ";" if sys.platform == "win32" else ":"
    cmd = [
        sys.executable, "-m", "PyInstaller",
        str(ROOT / entry_name),
        "--name", app_name,
        "--noconfirm",
        "--clean",
        "--console",
        *(["--onedir"] if onedir else ["--onefile"]),
        f"--add-data={PT_DST}{sep}platform-tools",
        "--collect-all=uiautomator2",
        "--collect-all=adbutils",
        "--collect-submodules=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=rich",
        "--exclude-module=tkinter",
        "--distpath", str(ROOT / "dist"),
        "--workpath", str(ROOT / "build"),
        "--specpath", str(ROOT),
    ]

    print("[build] " + " ".join(cmd))
    subprocess.run(cmd, cwd=ROOT, check=True)

    if sys.platform == "win32":
        exe = (
            ROOT / "dist" / app_name / f"{app_name}.exe"
            if onedir
            else ROOT / "dist" / f"{app_name}.exe"
        )
    else:
        exe = ROOT / "dist" / app_name / app_name
    if not exe.exists():
        raise SystemExit(f"[fail] 未找到产物: {exe}")
    size_mb = exe.stat().st_size / 1024 / 1024
    print(f"[ok] 打包完成: {exe}  ({size_mb:.1f} MB)")
    return exe


def main() -> int:
    parser = argparse.ArgumentParser(description="打包到梦空间自动化脚本为 exe")
    parser.add_argument("--onedir", action="store_true", help="生成目录版而非单文件版")
    parser.add_argument(
        "--variant",
        choices=[*VARIANTS, "all"],
        default="general",
        help="选择通用版、姓氏版、姓氏一分专用版，或全部构建",
    )
    args = parser.parse_args()
    variants = VARIANTS if args.variant == "all" else [args.variant]
    for variant in variants:
        build(onedir=args.onedir, variant=variant)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
