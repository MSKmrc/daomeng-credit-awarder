"""adb 可执行文件查找工具。

让「到梦空间」小工具在不同运行形态下都能找到 adb：
  1. 环境变量 ``ADBUTILS_ADB_PATH``（uiautomator2/adbutils 使用的变量）
  2. 随包/随 exe 携带的 ``platform-tools``（打包出来的小工具优先用自己的）
  3. 源码目录下的 ``daomeng/platform-tools``
  4. 环境变量 ``ANDROID_HOME`` / ``ANDROID_SDK_ROOT``
  5. 系统 ``PATH`` 里的 adb
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import List, Optional

ADB_NAME = "adb.exe" if sys.platform == "win32" else "adb"


def _exe_dir() -> Path:
    """返回可执行文件所在目录（源码运行时为入口脚本目录）。"""
    return Path(sys.argv[0]).resolve().parent


def _frozen_dir() -> Optional[Path]:
    """PyInstaller 单文件模式的临时解压目录。"""
    meipass = getattr(sys, "_MEIPASS", None)
    return Path(meipass) if meipass else None


def _package_dir() -> Path:
    """daomeng 包自身所在目录。"""
    return Path(__file__).resolve().parent


def candidate_paths() -> List[Path]:
    """按顺序返回所有可能的 adb 路径（不判断是否存在）。"""
    paths: List[Path] = []

    env_adb = os.environ.get("ADBUTILS_ADB_PATH")
    if env_adb:
        paths.append(Path(env_adb))

    frozen = _frozen_dir()
    if frozen:
        paths.append(frozen / "platform-tools" / ADB_NAME)
        paths.append(frozen / ADB_NAME)

    exe_dir = _exe_dir()
    paths.append(exe_dir / "platform-tools" / ADB_NAME)
    paths.append(exe_dir / ADB_NAME)

    pkg_dir = _package_dir()
    paths.append(pkg_dir / "platform-tools" / ADB_NAME)
    paths.append(pkg_dir.parent / "platform-tools" / ADB_NAME)

    for var in ("ANDROID_HOME", "ANDROID_SDK_ROOT"):
        root = os.environ.get(var)
        if root:
            paths.append(Path(root) / "platform-tools" / ADB_NAME)

    which = shutil.which("adb")
    if which:
        paths.append(Path(which))

    # 去重保序
    seen = set()
    unique: List[Path] = []
    for p in paths:
        key = str(p).lower()
        if key not in seen:
            seen.add(key)
            unique.append(p)
    return unique


def resolve_adb_path() -> Optional[Path]:
    """找到第一个真实存在的 adb 可执行文件。"""
    for path in candidate_paths():
        try:
            if path.is_file():
                return path
        except OSError:
            continue
    return None


def ensure_adb_path() -> Path:
    """确保能拿到 adb，并把结果写进 ``ADBUTILS_ADB_PATH``。

    Raises:
        FileNotFoundError: 所有候选位置都没有 adb 时抛出，错误信息列出搜索过的路径。
    """
    adb_path = resolve_adb_path()
    if adb_path is None:
        searched = "\n".join(f"  - {p}" for p in candidate_paths())
        raise FileNotFoundError(
            "没有找到 adb（platform-tools）。已尝试以下位置：\n"
            f"{searched}\n\n"
            "解决办法任选其一：\n"
            "  1. 使用自带的完整版小工具（同目录下应包含 platform-tools 文件夹）；\n"
            "  2. 把 Android SDK 的 platform-tools 文件夹放到程序所在目录；\n"
            "  3. 安装 Android SDK 并设置环境变量 ANDROID_HOME；\n"
            "  4. 设置环境变量 ADBUTILS_ADB_PATH 指向 adb.exe 的完整路径。"
        )

    resolved = str(adb_path.resolve())
    os.environ["ADBUTILS_ADB_PATH"] = resolved
    return adb_path
