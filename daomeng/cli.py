import argparse
import time
import json
import sys
from pandas import read_excel, notna
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from rich import print
from rich.prompt import Prompt

from .adb_helper import ensure_adb_path
from .automation_app import AutomationApp
from .config import AutomationConfig
from .device_controller import DeviceController
from .logging_utils import get_logger
from . import interactive


@dataclass
class CLIOptions:
    """命令行解析结果。"""

    device: Optional[str]
    pair: Optional[list[str]]
    additional_list: Optional[Path]
    console_log_level: str
    file_log_level: str
    file_dir: dict[float, Path]
    init_device: bool
    recovered: Optional[Path]
    accounts: Optional[dict[str, str]]
    yiban_mapping: Optional[Path]
    click_other_tab: bool
    account_usernames: Optional[list[str]]


def _save_to_json(config: AutomationConfig, config_path: Path) -> None:
    """保存任务配置。

    注意：出于安全考虑，密码**不会**写入磁盘，这里只保存账号名，
    恢复任务时会重新询问密码（或让用户通过 -u 再次提供账号表）。
    """
    path = config_path / "config.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump({
            "device": config.device,
            "additional_list": str(config.additional_list) if config.additional_list else None,
            "file_dir": [(str(path), pts) for pts, path in config.file_dir.items()] if config.file_dir else None,
            "account_usernames": list(config.accounts.keys()) if config.accounts else None,
            "yiban_mapping": str(config.yiban_mapping) if config.yiban_mapping else None,
            "click_other_tab": config.click_other_tab,
        }, f, ensure_ascii=False, indent=4)
        
        
def _load_from_json(config_path: Path) -> dict:
    path = config_path
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    data["additional_list"] = Path(data["additional_list"]) if data.get("additional_list") else None

    # 兼容旧版配置：早期版本会把账号密码明文写进 config.json，
    # 这里只继承账号名，密码一律不落盘、不复用。
    legacy_accounts = data.get("accounts")
    usernames = data.get("account_usernames")
    if not usernames and isinstance(legacy_accounts, dict):
        usernames = list(legacy_accounts.keys())
    data["account_usernames"] = usernames or None
    data["accounts"] = None

    data["yiban_mapping"] = Path(data["yiban_mapping"]) if data.get("yiban_mapping") else None
    data["click_other_tab"] = bool(data.get("click_other_tab", False))
    return data


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="daomeng-automation",
        description="到梦空间自动化脚本",
    )
    parser.add_argument("-d", help="无线调试 IP:端口，不传入则尝试有线连接。", metavar='IP:PORT')
    parser.add_argument("-p", help="无线调试配对 IP:端口 配对码。", metavar=('IP:PORT', 'CODE'), nargs=2)
    parser.add_argument("-a", help="额外添加名单。", type=Path, metavar='PATH')
    parser.add_argument(
        "-cl",
        default="WARNING",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        type=str.upper,
        help="控制台日志级别，默认 WARNING。"
    )
    parser.add_argument(
        "-fl",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        type=str.upper,
        help="文件日志级别，默认 INFO。",
    )
    parser.add_argument(
        "-f",
        nargs=2,
        action="append",
        metavar=("PATH", "POINTS"),
        help="传入 文件绝对路径 和 分值，可重复。录取名单分值设置为 0。",
    )
    parser.add_argument(
        "-i",
        action="store_true",
        help="初始化设备，第一次连接或链接异常时使用。",
    )
    parser.add_argument(
        "-r",
        type=Path,
        help="恢复上次未完成任务，传入上次任务结果文件夹绝对路径。注意：恢复模式下不允许修改 -f、-a、-yb 参数，但可以修改 -d 和 -u 参数。",
        metavar='PATH'
    )
    parser.add_argument(
        "-u",
        type=Path,
        help="传入 账号信息 文件绝对路径。",
        metavar='PATH'
    )
    parser.add_argument(
        "-yb",
        type=Path,
        help="传入 易班导出学生信息 CSV文件绝对路径，包含姓名、学号、易班ID三列。",
        metavar='PATH'
    )
    parser.add_argument(
        "--click-other-tab",
        action="store_true",
        help="在发分搜索前，先点击“已获得其他学分”tab再进行搜索。"
    )
    parser.add_argument(
        "--no-pause",
        action="store_true",
        help=argparse.SUPPRESS,  # exe 专用：退出前不暂停等待回车（批处理 wants this）
    )
    return parser


def _parse_args(argv: Optional[list[str]] = None) -> CLIOptions:
    parser = _build_parser()
    args = parser.parse_args(argv)
    file_dir: list[list[str]] = []
    file_dir_c: dict[float, Path] = {}
    device: Optional[str] = args.d
    additional_list: Optional[Path] = args.a
    accounts: Optional[dict[str, str]] = None
    yiban_mapping: Optional[Path] = None
    account_usernames: Optional[list[str]] = None
    click_other_tab: bool = args.click_other_tab
    
    if args.f is None and args.r is None:
        raise ValueError("必须指定 -r 或 -f 参数。")
    
    if args.f and args.r:
        raise ValueError("参数 -r 与 -f 不能同时使用。恢复模式下会自动使用原任务的配置。")
    
    if args.a and args.r:
        raise ValueError("参数 -a 与 -r 不能同时使用。恢复模式下会自动使用原任务的配置。")
    
    if args.yb and args.r:
        raise ValueError("参数 -yb 与 -r 不能同时使用。恢复模式下会自动使用原任务的配置。")

    if args.r:
        config_path = args.r / "config.json"
        if not config_path.exists():
            raise FileNotFoundError(f"指定的恢复路径无效，未找到配置文件: {config_path}")
        recovered_config = _load_from_json(config_path)
        
        # 恢复模式下只允许修改设备连接参数和账号信息
        device = args.d if args.d else recovered_config.get("device", "")
        additional_list = recovered_config.get("additional_list")
        file_dir = recovered_config.get("file_dir", [])
        click_other_tab = recovered_config.get("click_other_tab", False)
        account_usernames = recovered_config.get("account_usernames")
        
        # 账号信息：允许传入新的账号信息覆盖
        if args.u:
            # 将在后续处理
            pass
        
        # 易班映射：使用恢复的配置
        yiban_mapping_str = recovered_config.get("yiban_mapping")
        yiban_mapping = Path(yiban_mapping_str) if yiban_mapping_str else None
    
    file_dir = args.f if args.f else file_dir
    if file_dir:
        if len(file_dir) == 0:
            raise ValueError("文件路径列表不能为空。")

    for path_str, pts_str in file_dir:
        p = Path(path_str)
        pts = float(pts_str)
        if not p.exists():
            raise FileNotFoundError(f"指定的文件路径不存在: {p}")
        if p in file_dir_c.values():
            raise ValueError(f"检测到重复的文件路径: {p}")
        if pts in file_dir_c.keys():
            raise ValueError(f"检测到重复的分值: {pts}")
        file_dir_c[pts] = p
        
    if file_dir_c.get(0) is None:
        raise ValueError("必须指定分值为 0 的录取名单文件。")
    
    # 处理账号信息：新传入的覆盖恢复的（恢复模式和新任务模式都支持）
    if args.u:
        accounts_path = args.u
        if not accounts_path.exists():
            raise FileNotFoundError(f"指定的账号信息文件不存在: {accounts_path}")
        try:
            accounts_df = read_excel(accounts_path, usecols=["账号", "密码"], dtype=str, engine="openpyxl")
            accounts = {row["账号"]: row["密码"] for _, row in accounts_df.iterrows() if notna(row["账号"]) and notna(row["密码"])}
        except Exception as e:
            raise ValueError(f"读取账号信息文件失败: {e}")
    
    # 处理易班ID映射：仅在非恢复模式下处理新传入的映射
    if args.yb and not args.r:
        yiban_mapping = args.yb
    
    if yiban_mapping and not yiban_mapping.exists():
        raise FileNotFoundError(f"指定的易班ID映射文件不存在: {yiban_mapping}")
    
    if args.p and len(args.p) != 2:
        raise ValueError("参数 -p 需要传入两个值：IP:端口 和 配对码。")
            
        
    return CLIOptions(
        device = device if args.r else args.d,
        pair = args.p,
        additional_list = additional_list if args.r else args.a,
        console_log_level = args.cl,
        file_log_level = args.fl,
        file_dir = file_dir_c,
        init_device = args.i,
        recovered = args.r,
        accounts = accounts if accounts else None,
        yiban_mapping = yiban_mapping if yiban_mapping else None,
        click_other_tab = click_other_tab,
        account_usernames = account_usernames if not accounts else None,
    )


def _read_activity_name(enrollment: Path) -> str:
    """从录取名单里读活动名称。

    取「活动名称」列中第一个非空值，避免因为活动名只写在首行（或前几行有空单元格）
    而读到 NaN。读不到时给出可操作的报错。
    """
    try:
        series = read_excel(
            enrollment,
            usecols=lambda col: str(col).strip() == "活动名称",
            dtype=str,
            engine="openpyxl",
        ).iloc[:, 0]
    except Exception:  # 没有该列时 pandas 会抛 ValueError 或返回空表导致 IndexError
        series = None
    if series is None or len(series) == 0:
        raise ValueError(
            f"录取名单缺少「活动名称」列：{enrollment}\n"
            "  录取名单（0 分那份）必须有表头「活动名称」，并在该列填上本次活动的完整名称；\n"
            "  脚本靠它在手机 APP 里搜索活动。现有列名："
            f"{list(read_excel(enrollment, nrows=0, engine='openpyxl').columns)}"
        )

    values = [v.strip() for v in series.tolist() if isinstance(v, str) and v.strip() and v.strip().lower() != "nan"]
    if not values:
        raise ValueError(
            f"录取名单的「活动名称」列是空的：{enrollment}\n"
            "  请在该列至少填一个活动名称（要和 APP 里显示的活动名完全一致）。"
        )

    name = values[0]
    if len(set(values)) > 1:
        print(f"[yellow]警告：录取名单里出现多个活动名称 {sorted(set(values))}，将使用第一个：{name}[/yellow]")
    return name


def _probe_screen(argv: Optional[list[str]]) -> int:
    """把手机当前界面的控件树导出来，用于发现「批量勾选 / 颁发N人」这类更快的入口。

    用法：
      --probe                      只导出当前屏幕
      --probe -f 录取名单.xlsx 0 -f 1分名单.xlsx 1.0
                                   自动导航到第一个非零分值的发分页再导出
    """
    from xml.etree import ElementTree as ET

    src = list(argv) if argv is not None else sys.argv[1:]
    rest = [a for a in src if a != "--probe"]

    ensure_adb_path()
    logger = get_logger(console_log_level="INFO", log_file=None)

    base_dir = Path(sys.argv[0]).resolve().parent
    out_dir = base_dir / f"probe-{time.strftime('%Y%m%d-%H%M%S')}"
    out_dir.mkdir(parents=True, exist_ok=True)

    if "-f" in rest:
        options = _parse_args(rest)
        activity_name = _read_activity_name(options.file_dir.get(0))
        scores = sorted(s for s in options.file_dir if s > 0)
        if not scores:
            print("[red]--probe 需要至少一份分值大于 0 的发分名单才能自动导航。[/red]")
            return 1
        target_score = scores[0]
        print(f"[bold]将自动导航到 {target_score} 分发放界面后导出控件树[/bold]")
        config = AutomationConfig(
            device=options.device,
            pair=options.pair,
            additional_list=None,
            console_log_level="INFO",
            file_log_level="DEBUG",
            file_dir=options.file_dir,
            output_dir=out_dir,
            init_device=options.init_device,
            accounts=options.accounts,
            activity_name=activity_name,
            yiban_mapping=None,
            click_other_tab=options.click_other_tab,
        )
        controller = DeviceController(config, logger)
        controller.connect()
        controller.navigation(target_score)
        device = controller.device
    else:
        import uiautomator2 as u2

        target_score = None
        device_arg = None
        if "-d" in rest:
            idx = rest.index("-d")
            if idx + 1 < len(rest):
                device_arg = rest[idx + 1]
        print("[bold]直接导出手机当前屏幕的控件树（请先把手机停在要检查的页面上）[/bold]")
        device = u2.connect(device_arg) if device_arg else u2.connect()
        device.settings["wait_timeout"] = 15.0

    xml = device.dump_hierarchy()
    xml_file = out_dir / "hierarchy.xml"
    xml_file.write_text(xml, encoding="utf-8")

    KEYWORDS = ("check", "select", "radio", "cb_", "cbx", "btn", "step", "screen",
                "title", "score", "next", "search", "confirm", "all")
    lines: list[str] = []
    interesting: list[str] = []
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as exc:
        print(f"[red]控件树解析失败: {exc}（原始 xml 已保存到 {xml_file}）[/red]")
        return 1

    for node in root.iter("node"):
        res = node.attrib.get("resource-id", "")
        cls = node.attrib.get("class", "")
        text = (node.attrib.get("text") or "").strip()
        desc = (node.attrib.get("content-desc") or "").strip()
        bounds = node.attrib.get("bounds", "")
        flags = "".join([
            "K" if node.attrib.get("clickable") == "true" else "-",
            "X" if node.attrib.get("checkable") == "true" else "-",
            "V" if node.attrib.get("checked") == "true" else "-",
            "S" if node.attrib.get("selected") == "true" else "-",
            "E" if node.attrib.get("enabled") == "true" else "-",
        ])
        row = f"{flags} {bounds:<24} {cls.split('.')[-1]:<22} {res:<52} text={text!r} desc={desc!r}"
        lines.append(row)
        low = (res + cls + desc).lower()
        if any(k in low for k in KEYWORDS) or text:
            interesting.append(row)

    report = out_dir / "probe.txt"
    report.write_text(
        "K=clickable X=checkable V=checked S=selected E=enabled\n"
        f"{'='*100}\n【重点嫌疑控件】\n{'='*100}\n"
        + "\n".join(interesting)
        + f"\n\n{'='*100}\n【全部控件】\n{'='*100}\n"
        + "\n".join(lines),
        encoding="utf-8",
    )

    print(f"[green]导出完成：{report}[/green]")
    print(f"[dim]原始控件树：{xml_file}（共 {len(lines)} 个节点）[/dim]")
    print("\n[bold]下面是可能被用于批量操作的控件（K=可点 X=可选中 V=已选中）：[/bold]")
    for row in interesting[:80]:
        print("  " + row)
    if len(interesting) > 80:
        print(f"  ... 其余 {len(interesting) - 80} 条见 probe.txt")
    print("\n[dim]把这个 probe.txt 发给开发者，就能判断 APP 是否支持「勾选多人 → 一次颁发」的批量模式。[/dim]")
    return 0


def run_cli(argv: Optional[list[str]] = None) -> int:
    print("[bold cyan]到梦空间自动化脚本[/bold cyan]")
    print("="*60)

    raw = argv if argv is not None else sys.argv[1:]
    if "--probe" in raw:
        return _probe_screen(argv)
    
    # 没有任何参数（一般是双击 exe）时进入交互式向导，避免“一闪而过”
    if argv is None and len(sys.argv) == 1:
        try:
            argv = interactive.build_args_interactively()
        except KeyboardInterrupt:
            print("\n[bold yellow]已取消。[/bold yellow]")
            return 0
        except Exception as e:
            print(f"[bold red]向导中断:[/bold red] {e}")
            return 1
        interactive.wizard_used = True
        if argv is None:
            return 0
        print("="*60)
    
    try:
        options = _parse_args(argv)
    except Exception as e:
        print(f"[bold red]命令行参数解析失败:[/bold red] {e}")
        return 1
    
    activity_name = _read_activity_name(options.file_dir.get(0))
    print(f"[bold yellow]活动名称:[/bold yellow] [bold green]{activity_name}[/bold green]")

    timestamp = time.strftime("%Y%m%d-%H%M%S")
    base_dir = Path(sys.argv[0]).resolve().parent
    if options.recovered is None:
        output_dir = (
            base_dir
            / f"{activity_name.replace('“', '').replace('”', '').replace('"', '')}-{timestamp}"
        )
        output_dir.mkdir(parents=True, exist_ok=True)
        print(f"[bold yellow]输出目录:[/bold yellow] {output_dir}")
    else:
        output_dir = options.recovered
        print(f"[bold yellow]恢复任务:[/bold yellow] {output_dir}")
    log_file = output_dir / "automation.log"
    
    logger = get_logger(
        console_log_level=options.console_log_level,
        file_log_level=options.file_log_level,
        log_file=log_file
    )
    
    logger.info("="*60)
    logger.info("启动到梦空间自动化脚本")
    logger.info(f"活动名称: {activity_name}")
    logger.info(f"输出目录: {output_dir}")
    logger.info(f"日志文件: {log_file}")
    logger.info(f"控制台日志级别: {options.console_log_level}")
    logger.info(f"文件日志级别: {options.file_log_level}")
    logger.info("="*60)
    
    # 恢复模式下若曾用过自动登录，密码不落盘，这里重新询问
    if not options.accounts and options.account_usernames:
        restored: dict[str, str] = {}
        for username in options.account_usernames:
            password = Prompt.ask(
                f"恢复任务：请输入账号 [bold cyan]{username}[/bold cyan] 的密码（直接回车跳过自动登录）",
                password=True,
                default="",
            )
            if password:
                restored[username] = password
        options.accounts = restored or None

    config = AutomationConfig(
        device=options.device,
        pair=options.pair,
        additional_list=options.additional_list,
        console_log_level=options.console_log_level,
        file_log_level=options.file_log_level,
        file_dir=options.file_dir,
        output_dir=output_dir,
        init_device=options.init_device,
        accounts=options.accounts,
        activity_name=activity_name,
        yiban_mapping=options.yiban_mapping,
        click_other_tab=options.click_other_tab,
    )
    
    _save_to_json(config, output_dir)
    logger.info("配置信息已保存到 config.json")
    logger.debug(f"当前配置: {config}")

    # ADB 查找顺序：环境变量 → 随包携带的 platform-tools → 源码/exe 同级目录
    #               → ANDROID_HOME/ANDROID_SDK_ROOT → 系统 PATH
    adb_path = ensure_adb_path()
    logger.info(f"使用 ADB: {adb_path}")

    app = AutomationApp(
        config=config,
        logger=logger,
    )
    
    print("[bold green]开始执行自动化任务...[/bold green]")
    print("="*60)

    return app.run()
