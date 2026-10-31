"""双击 exe 时的交互式向导。

没有命令行参数时（通常是双击运行）引导用户一步步填好名单和连接参数，
最后拼成等价的命令行参数列表，复用原有的 CLI 逻辑。
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional

from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt

console = Console()

EXTS = (".xlsx", ".xlsm", ".xls")

# 本次运行是否使用了向导模式（exe 据此决定退出前是否暂停）
wizard_used = False


def _clean(raw: str) -> str:
    """去掉拖拽文件时常见的前后引号和空白。"""
    return raw.strip().strip('"').strip("'").strip()


def _scan_candidates() -> List[Path]:
    """扫描「当前目录」和「桌面」里的表格文件，方便直接选编号。"""
    bases = [Path.cwd()]
    desktop = Path.home() / "Desktop"
    if desktop.exists():
        bases.append(desktop)
    found: List[Path] = []
    for base in bases:
        try:
            for p in sorted(base.iterdir()):
                if p.suffix.lower() in EXTS and p not in found:
                    found.append(p)
        except OSError:
            continue
    return found[:20]


def pick_file(title: str, *, required: bool = True) -> Optional[Path]:
    """让用户选择一个文件，支持直接输入路径或把文件拖进窗口。"""
    candidates = _scan_candidates()
    if candidates:
        console.print(f"\n[bold cyan]{title}[/bold cyan]\n[dim]检测到这些表格，直接输编号即可：[/dim]")
        for idx, p in enumerate(candidates, 1):
            console.print(f"  [bold]{idx}[/bold]. {p.name}  [dim]({p.parent})[/dim]")
        console.print("  [bold]0[/bold]. 手动输入完整路径（或把文件拖进窗口后回车）")

    while True:
        raw = Prompt.ask(f"[cyan]{title}[/cyan]{'' if required else '（可直接回车跳过）'}")
        value = _clean(raw)
        if not value:
            if not required:
                return None
            console.print("[yellow]这项是必填的。[/yellow]")
            continue

        if candidates and value.isdigit():
            num = int(value)
            if 1 <= num <= len(candidates):
                return candidates[num - 1]
            if num == 0:
                continue
            console.print("[yellow]编号不在列表里，请重新输入。[/yellow]")
            continue

        path = Path(value)
        if path.is_file():
            return path
        console.print(f"[red]文件不存在:[/red] {path}")


def ask_points() -> Optional[str]:
    """询问分值，返回分值字符串；用户输入 q 表示结束添加。"""
    while True:
        value = _clean(Prompt.ask("[cyan]这份名单发几分？[/cyan]（如 1.0，输入 q 结束添加名单）", default=""))
        if value.lower() == "q":
            return None
        if not value:
            continue
        try:
            float(value)
            return value
        except ValueError:
            console.print("[yellow]请输入数字分值，例如 1.0[/yellow]")


def _row_count(path: Path) -> Optional[int]:
    """读一下名单有多少行；读不出来就返回 None，不影响主流程。"""
    try:
        from pandas import read_excel  # 局部导入，避免向导阶段就依赖 pandas

        return len(read_excel(path, engine="openpyxl", dtype=str, usecols=lambda c: c in {"学号", "姓名"}))
    except Exception:
        return None


def _print_summary(args: List[str]) -> None:
    """把「哪份名单发几分、多少人」列出来给用户最后确认。"""
    rows: List[tuple] = []
    i = 0
    while i < len(args):
        if args[i] == "-f":
            rows.append((Path(args[i + 1]), args[i + 2]))
            i += 3
        else:
            i += 1
    if not rows:
        return
    for path, points in rows:
        cnt = _row_count(path)
        cnt_text = f"{cnt} 人" if cnt is not None else "（人数读取失败）"
        role = "[dim]基准，不发分[/dim]" if float(points) == 0 else f"[bold green]{points} 分[/bold green]"
        console.print(f"  · {path.name}  →  {role}   [cyan]{cnt_text}[/cyan]  [dim]{path.parent}[/dim]")
    console.print()


def build_args_interactively() -> Optional[List[str]]:
    """跑完整向导，返回等价的命令行参数列表；用户中途退出返回 None。"""
    console.print(Panel.fit(
        "[bold]到梦空间学分助手[/bold]\n"
        "[dim]按提示一步步填写即可，找不到文件时可以直接把表格拖进这个窗口再回车。\n"
        "开始前请确认：手机已开启「开发者选项 - USB 调试」，并且登录了有学分管理权限的账号。[/dim]",
        border_style="cyan",
    ))

    args: List[str] = []

    # ① 录取名单
    console.print(
        "\n[bold]① 录取名单[/bold]（发 0 分的基准名单，必须包含 [cyan]活动名称 / 学号 / 姓名[/cyan] 三列）\n"
        "[dim]这份不发分，只用来核对姓名，并且脚本从它的「活动名称」列读取本次要去 APP 里搜索的活动名。[/dim]"
    )
    enrollment = pick_file("录取名单", required=True)
    if enrollment is None:  # pragma: no cover
        return None
    args += ["-f", str(enrollment), "0"]

    # ② 发分名单（可多份）
    console.print(
        "\n[bold]② 发分名单[/bold]（可以一份一份加，每份对应一个分值，需要 [cyan]学号 / 姓名[/cyan] 两列）\n"
        "[dim]例如：先选「1分.xlsx」→ 输入 1.0；再选「2分.xlsx」→ 输入 2.0；加完后直接回车结束。\n"
        "文件名不会被解析，分数全看你这里填的数字。[/dim]"
    )
    while True:
        score_file = pick_file("下一份发分名单（直接回车＝加完了）", required=False)
        if score_file is None:
            break
        points = ask_points()
        if points is None:
            break
        args += ["-f", str(score_file), points]
        console.print(f"[green]已添加：{score_file.name} → {points} 分[/green]")

    if len(args) == 3:
        raise ValueError("至少需要一份分值大于 0 的发分名单（录取名单分值为 0，本身不发分）")

    # ③ 补充名单 / 账号表
    extra = pick_file("③ 补充名单（可选，需含 学号 / 姓名 / 分值）", required=False)
    if extra:
        args += ["-a", str(extra)]

    accounts = pick_file("④ 账号表（可选，掉登录时自动重登，需含 账号 / 密码）", required=False)
    if accounts:
        args += ["-u", str(accounts)]

    yiban = pick_file("⑤ 易班名单 CSV（可选，名单没学号列时才需要）", required=False)
    if yiban:
        args += ["-yb", str(yiban)]

    # ⑥ 设备连接
    console.print("\n[bold]⑥ 设备连接[/bold]")
    console.print("[dim]  用数据线：直接回车\n"
                  "  用无线调试：填手机的 IP:端口，例如 192.168.1.100:5555[/dim]")
    addr = _clean(Prompt.ask("连接地址", default=""))
    if addr:
        args += ["-d", addr]
        pair_addr = _clean(Prompt.ask(
            "是否需要先配对？（手机「无线调试」页显示的配对 IP:端口，不需要就留空）", default=""))
        if pair_addr:
            pair_code = _clean(Prompt.ask("配对码（手机上那 6 位数字）"))
            if pair_code:
                args += ["-p", pair_addr, pair_code]

    # ⑦ 开关项
    if Prompt.ask("\n⑦ 这台手机是不是第一次用本工具？（是＝初始化依赖）",
                  choices=["y", "n"], default="n") == "y":
        args.append("-i")

    if Prompt.ask("⑧ 发分前先点「已获得其他学分」页签？",
                  choices=["y", "n"], default="n") == "y":
        args.append("--click-other-tab")

    level = Prompt.ask("⑨ 控制台日志详细程度", choices=["WARNING", "INFO", "DEBUG"], default="WARNING")
    args += ["-cl", level]

    # 确认
    console.print("\n[bold]即将按以下配置运行：[/bold]")
    _print_summary(args)
    console.print("[dim]  " + " ".join(f'"{a}"' if " " in a else a for a in args) + "[/dim]")
    if Prompt.ask("确认开始？", choices=["y", "n"], default="y") != "y":
        console.print("[yellow]已取消。[/yellow]")
        return None
    return args
