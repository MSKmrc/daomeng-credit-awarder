import collections
import contextlib
import datetime
import importlib
import json
import os
import pathlib
import re
import subprocess
import sys
import time
import xml.etree.ElementTree as ET


def _fix_windows_encoding() -> None:
    if sys.platform != "win32":
        return
    try:
        subprocess.run(
            ["cmd", "/c", "chcp", "65001", ">NUL"],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        for stream in (sys.stdin, sys.stdout, sys.stderr):
            try:
                if stream is not None and hasattr(stream, "reconfigure"):
                    stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
    except Exception:
        pass


_fix_windows_encoding()

from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt

import daomeng.interactive as interactive
from daomeng.cli import run_cli

console = Console()


def _show_operation_guide() -> None:
    console.print(
        Panel(
            "[bold]请按以下顺序准备：[/bold]\n\n"
            "[cyan]1. 手机[/cyan]\n"
            "  保持解锁并连接电脑；开启 USB 调试，手机上出现授权提示时选择允许。\n"
            "  确认到梦空间当前登录的是 [yellow]有活动管理权限的账号[/yellow]。\n\n"
            "[cyan]2. 录取名单（必填，0 分）[/cyan]\n"
            "  选择本次活动的报名/录取表，必须包含"
            " [green]活动名称、学号、姓名[/green]。\n"
            "  程序只把它作为核对基准，[bold]不会给这份名单发分[/bold]。\n\n"
            "[cyan]3. 发分名单（必填）[/cyan]\n"
            "  选择实际需要发分的人员表，必须包含 [green]学号、姓名[/green]。\n"
            "  文件名不会决定分值；在下一步手工输入本次分值的数字，例如 "
            "[bold]2.0[/bold]。\n\n"
            "[cyan]4. 可选材料[/cyan]\n"
            "  账号表、易班 ID 映射、附加名单若没有就留空；"
            "当前已登录时通常不需要账号表。\n\n"
            "[cyan]5. 开始前核对[/cyan]\n"
            "  最后会显示文件、人数和分值摘要。必须确认活动名称、"
            "[bold]分值和名单数量[/bold]都正确。\n"
            "  学分提交后不能由本程序撤销，结果不明确时只会标记“待核对”。\n\n"
            "[cyan]6. 结果保存位置[/cyan]\n"
            "  程序会在当前 EXE 所在目录自动创建"
            " [bold]活动名称-日期时间[/bold] 文件夹。\n"
            "  里面包含 automation.log 和同名 Excel 发放结果表。\n"
            "  即使人员没有搜索到，只要流程正常结束，也会生成结果表和日志。\n\n"
            "[bold yellow]本软件由南华大学易班学生工作站制作[/bold yellow]",
            title="到梦空间学分助手 · 使用前必看",
            border_style="cyan",
        )
    )


_original_build_args_interactively = interactive.build_args_interactively


def _guided_build_args_interactively():
    _show_operation_guide()
    return _original_build_args_interactively()


interactive.build_args_interactively = _guided_build_args_interactively


def _normalized_cell(value) -> str:
    return str(value).strip() if value is not None else ""


def _find_column(headers, name):
    normalized = [_normalized_cell(value).replace("\ufeff", "") for value in headers]
    if name not in normalized:
        raise ValueError(f"缺少列“{name}”。当前表头：{', '.join(normalized)}")
    return normalized.index(name)


def _find_optional_column(headers, name):
    normalized = [_normalized_cell(value).replace("\ufeff", "") for value in headers]
    return normalized.index(name) if name in normalized else None


def _read_first_sheet(path, role):
    import openpyxl

    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook[workbook.sheetnames[0]]
        iterator = sheet.iter_rows(values_only=True)
        headers = next(iterator)
        rows = [
            (excel_row, record)
            for excel_row, record in enumerate(iterator, 2)
            if any(_normalized_cell(value) for value in record)
        ]
        return list(headers), rows
    except Exception as exc:
        raise ValueError(f"{role}读取失败：{exc}") from exc
    finally:
        workbook.close()


def _prepare_surname_batches(enrollment_path, score_path):
    enrollment_headers, enrollment_rows = _read_first_sheet(
        enrollment_path, "录取名单"
    )
    score_headers, score_rows = _read_first_sheet(score_path, "发分名单")
    enrollment_id_col = _find_column(enrollment_headers, "学号")
    enrollment_name_col = _find_column(enrollment_headers, "姓名")
    enrollment_activity_col = _find_column(enrollment_headers, "活动名称")
    enrollment_yiban_col = _find_optional_column(enrollment_headers, "易班ID")
    enrollment_college_col = _find_optional_column(enrollment_headers, "学院")
    enrollment_class_col = _find_optional_column(enrollment_headers, "班级")
    score_id_col = _find_column(score_headers, "学号")
    score_name_col = _find_column(score_headers, "姓名")

    enrollment_names = {}
    enrollment_name_counts = collections.Counter()
    enrollment_records = []
    activity_name = ""
    for _, row in enrollment_rows:
        student_id = _normalized_cell(row[enrollment_id_col])
        name = _normalized_cell(row[enrollment_name_col])
        row_activity = _normalized_cell(row[enrollment_activity_col])
        if not activity_name and row_activity:
            activity_name = row_activity
        if name:
            enrollment_name_counts[name] += 1
        if student_id and name:
            enrollment_names.setdefault(student_id, set()).add(name)
        enrollment_records.append(
            {
                "excel_row": _,
                "name": name,
                "student_id": student_id,
                "yiban_id": (
                    _normalized_cell(row[enrollment_yiban_col])
                    if enrollment_yiban_col is not None
                    else ""
                ),
                "college": (
                    _normalized_cell(row[enrollment_college_col])
                    if enrollment_college_col is not None
                    else ""
                ),
                "class_name": (
                    _normalized_cell(row[enrollment_class_col])
                    if enrollment_class_col is not None
                    else ""
                ),
            }
        )

    pending = []
    deduplicated = []
    by_id = collections.OrderedDict()
    for excel_row, row in score_rows:
        student_id = _normalized_cell(row[score_id_col])
        name = _normalized_cell(row[score_name_col])
        if not re.fullmatch(r"\d{11}", student_id):
            pending.append(
                {
                    "excel_row": excel_row,
                    "student_id": student_id,
                    "name": name,
                    "surname": name[:1],
                    "reason": "学号不是11位数字",
                    "status": "待核查",
                }
            )
            continue
        by_id.setdefault(student_id, []).append((excel_row, name))

    unique_scores = []
    for student_id, records in by_id.items():
        names = {record_name for _, record_name in records}
        if len(names) > 1:
            for excel_row, name in records:
                pending.append(
                    {
                        "excel_row": excel_row,
                        "student_id": student_id,
                        "name": name,
                        "surname": name[:1],
                        "reason": "同一学号对应多个姓名",
                        "status": "待核查",
                    }
                )
            continue

        excel_row, name = records[-1]
        unique_scores.append((excel_row, student_id, name))
        for duplicate_row, _ in records[:-1]:
            deduplicated.append(
                {
                    "excel_row": duplicate_row,
                    "student_id": student_id,
                    "name": name,
                    "reason": f"完全重复，保留最后第 {excel_row} 行",
                }
            )

    score_name_counts = collections.Counter(
        name for _, _, name in unique_scores if name
    )
    accepted = []
    accepted_ids = set()
    for excel_row, student_id, name in unique_scores:
        reason = None
        if student_id not in enrollment_names:
            reason = "学号不在录取名单"
        elif name not in enrollment_names[student_id]:
            reason = "姓名与录取名单不一致"
        elif score_name_counts[name] > 1:
            reason = "发分名单中存在重名"

        record = {
            "excel_row": excel_row,
            "student_id": student_id,
            "name": name,
            "surname": name[:1],
            "reason": reason or "",
            "status": "待核查" if reason else "待执行",
        }
        if reason:
            pending.append(record)
        else:
            accepted_ids.add(student_id)
            accepted.append(record)

    grouped = collections.OrderedDict()
    accepted.sort(
        key=lambda item: (item["surname"], item["name"], item["student_id"])
    )
    for record in accepted:
        grouped.setdefault(record["surname"], []).append(record)
    if not activity_name:
        raise ValueError("录取名单的“活动名称”列为空")

    duplicate_names = {
        name for name, count in score_name_counts.items() if count > 1
    }
    score_rows_by_name = collections.defaultdict(list)
    for excel_row, student_id, name in unique_scores:
        score_rows_by_name[name].append((excel_row, student_id))
    duplicate_report = []
    for record in enrollment_records:
        if record["name"] not in duplicate_names:
            continue
        score_matches = score_rows_by_name.get(record["name"], [])
        duplicate_report.append(
            {
                **record,
                "name_count": enrollment_name_counts[record["name"]],
                "score_excel_rows": ",".join(
                    str(row) for row, _ in score_matches
                ),
                "score_student_ids": ",".join(
                    student_id for _, student_id in score_matches
                ),
            }
        )
    duplicate_report.sort(
        key=lambda item: (item["name"], item["student_id"])
    )
    return (
        accepted,
        pending,
        deduplicated,
        grouped,
        activity_name,
        duplicate_report,
    )


def _batch_output_root():
    override = os.environ.get("DAOMENG_BATCH_ROOT")
    if override:
        root = pathlib.Path(override)
    else:
        root = pathlib.Path.home() / "Desktop" / "到梦空间发分结果"
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    output = root / f"按姓氏批次-{stamp}"
    output.mkdir(parents=True, exist_ok=True)
    return output


def _write_batch_reports(
    output_dir, accepted, pending, deduplicated, duplicate_report, points
):
    import openpyxl

    pending_path = output_dir / "01-待核查-重名与异常.xlsx"
    score_label = f"{points:g}分"
    plan_path = output_dir / f"02-按姓氏排列的{score_label}名单.xlsx"
    dedup_path = output_dir / "00-自动去重记录.xlsx"
    duplicate_path = (
        output_dir / "03-重名人员对照表.xlsx"
        if duplicate_report
        else None
    )

    dedup_workbook = openpyxl.Workbook()
    dedup_sheet = dedup_workbook.active
    dedup_sheet.title = "自动去重"
    dedup_sheet.append(["原名单位置", "姓名", "学号", "处理说明"])
    for record in deduplicated:
        dedup_sheet.append(
            [
                record["excel_row"],
                record["name"],
                record["student_id"],
                record["reason"],
            ]
        )
    dedup_workbook.save(dedup_path)
    dedup_workbook.close()

    if duplicate_path is not None:
        duplicate_workbook = openpyxl.Workbook()
        duplicate_sheet = duplicate_workbook.active
        duplicate_sheet.title = "重名人员"
        duplicate_sheet.append(
            [
                "姓名",
                "班级",
                "学号",
                "易班ID",
                "学院",
                "同名人数量",
                "2分名单行号",
                "2分名单学号",
            ]
        )
        for record in duplicate_report:
            duplicate_sheet.append(
                [
                    record["name"],
                    record["class_name"],
                    record["student_id"],
                    record["yiban_id"],
                    record["college"],
                    record["name_count"],
                    record["score_excel_rows"],
                    record["score_student_ids"],
                ]
            )
        duplicate_workbook.save(duplicate_path)
        duplicate_workbook.close()

    pending_workbook = openpyxl.Workbook()
    pending_sheet = pending_workbook.active
    pending_sheet.title = "待核查"
    pending_sheet.append(["原名单位置", "姓名", "学号", "姓", "原因"])
    for record in pending:
        pending_sheet.append(
            [
                record["excel_row"],
                record["name"],
                record["student_id"],
                record["surname"],
                record["reason"],
            ]
        )
    pending_workbook.save(pending_path)
    pending_workbook.close()

    plan_workbook = openpyxl.Workbook()
    plan_sheet = plan_workbook.active
    plan_sheet.title = "正常计划"
    plan_sheet.append(["姓", "姓名", "学号", "原名单行号", "分值", "状态"])
    for record in accepted:
        plan_sheet.append(
            [
                record["surname"],
                record["name"],
                record["student_id"],
                record["excel_row"],
                points,
                record["status"],
            ]
        )
    plan_workbook.save(plan_path)
    plan_workbook.close()

    summary = {
        "accepted": len(accepted),
        "pending": len(pending),
        "deduplicated": len(deduplicated),
        "duplicate_people": len(duplicate_report),
        "duplicate_names": len(
            {record["name"] for record in duplicate_report}
        ),
        "surname_groups": len({record["surname"] for record in accepted}),
        "points": points,
        "pending_reasons": dict(
            collections.Counter(record["reason"] for record in pending)
        ),
    }
    (output_dir / "04-批次摘要.txt").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return pending_path, plan_path, duplicate_path, summary


def _write_batch_results(output_dir, accepted, pending, points, run_notes):
    import openpyxl

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "发放结果"
    sheet.append(["姓", "姓名", "学号", "原名单行号", "分值", "状态"])
    for record in accepted:
        sheet.append(
            [
                record["surname"],
                record["name"],
                record["student_id"],
                record["excel_row"],
                points,
                record["status"],
            ]
        )
    pending_sheet = workbook.create_sheet("待核查")
    pending_sheet.append(["原名单位置", "姓名", "学号", "姓", "原因", "状态"])
    uncertain = [
        record
        for record in accepted
        if "待核查" in record["status"] or "失败" in record["status"]
    ]
    for record in pending + uncertain:
        reason = record.get("reason") or record["status"]
        pending_sheet.append(
            [
                record["excel_row"],
                record["name"],
                record["student_id"],
                record["surname"],
                reason,
                record["status"],
            ]
        )
    notes_sheet = workbook.create_sheet("运行说明")
    notes_sheet.append(["说明"])
    for note in run_notes:
        notes_sheet.append([note])
    result_path = output_dir / "05-最终汇总.xlsx"
    workbook.save(result_path)
    workbook.close()
    return result_path


def _ensure_system_keyboard(device):
    try:
        output = device.shell(["ime", "list", "-s"]).output
        choices = [
            line.strip()
            for line in (output or "").splitlines()
            if line.strip() and "uiautomator" not in line.lower()
        ]
        if choices:
            device.shell(["ime", "set", choices[0]])
            return choices[0]
    except Exception:
        pass
    return None


def _click_keyboard_search(device):
    width, height = device.window_size()
    device.click(int(width * 0.895), int(height * 0.930))
    time.sleep(1.0)
    try:
        device.press("back")
    except Exception:
        pass
    time.sleep(0.4)


def _show_system_keyboard(device, element):
    match = re.fullmatch(
        r"\[(\d+),(\d+)\]\[(\d+),(\d+)\]",
        str(element.info.get("bounds", "")),
    )
    bounds_info = tuple(map(int, match.groups())) if match else None
    if bounds_info:
        left, top, right, bottom = bounds_info
        device.click((left + right) // 2, (top + bottom) // 2)
    else:
        element.click()
    deadline = time.monotonic() + 4
    while time.monotonic() < deadline:
        try:
            state = device.shell("dumpsys input_method").output or ""
            if "mInputShown=true" in state:
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


def _parse_bounds(node):
    match = re.fullmatch(
        r"\[(\d+),(\d+)\]\[(\d+),(\d+)\]", node.attrib.get("bounds", "")
    )
    return tuple(map(int, match.groups())) if match else None


def _batch_snapshot(device):
    root = ET.fromstring(device.dump_hierarchy())
    parents = {child: parent for parent in root.iter() for child in parent}
    return root, parents


def _clickable_ancestor(node, parents):
    current = node
    while current is not None:
        if current.attrib.get("clickable") == "true":
            return current
        current = parents.get(current)
    return None


def _selected_count(device, package_name):
    try:
        text = device(
            resourceId=f"{package_name}:id/tv_step"
        ).get_text()
        match = re.search(r"(\d+)", text or "")
        return int(match.group(1)) if match else 0
    except Exception:
        return 0


def _search_box_for_surname(device, surname, keyboard_ime=None):
    box = device(resourceId=f"com.jingcai.apps.qualitydev:id/text_searchcontent")
    deadline = time.monotonic() + 10
    while not box.exists and time.monotonic() < deadline:
        time.sleep(0.25)
    if not box.exists:
        raise RuntimeError("人员搜索框未出现")
    _ensure_system_keyboard(device)
    time.sleep(0.5)
    if not _show_system_keyboard(device, box):
        raise RuntimeError("系统输入法键盘没有弹出")
    box.set_text(surname)
    time.sleep(0.5)
    current = box.get_text().strip()
    if surname not in current:
        raise RuntimeError(f"姓氏未写入学号搜索框，当前='{current}'")

    _click_keyboard_search(device)


def _visible_target_rows(device, target_names):
    root, parents = _batch_snapshot(device)
    try:
        width, _ = device.window_size()
    except Exception:
        width = 10000
    found = {}
    visible_names = set()
    for node in root.iter("node"):
        name = node.attrib.get("text", "").strip()
        resource_id = node.attrib.get("resource-id", "")
        if (
            resource_id.endswith(":id/title")
            and name
            and node.attrib.get("clickable") != "true"
        ):
            node_bounds = _parse_bounds(node)
            if (
                node_bounds
                and node_bounds[0] >= 0
                and node_bounds[2] <= width
                and node_bounds[1] >= 300
                and node_bounds[3] <= 2300
            ):
                visible_names.add(name)
        if name not in target_names:
            continue
        node_bounds = _parse_bounds(node)
        if not node_bounds or node_bounds[0] < 0 or node_bounds[2] > width:
            continue
        row = _clickable_ancestor(node, parents)
        if row is None:
            continue
        check_bounds = None
        for child in row.iter("node"):
            resource_id = child.attrib.get("resource-id", "")
            if resource_id.endswith(":id/ll_imgcheck"):
                check_bounds = _parse_bounds(child)
                break
        if check_bounds is None:
            title_center = sum(_parse_bounds(node)[1::2]) // 2
            candidates = []
            for child in row.iter("node"):
                if child.attrib.get("resource-id", "").endswith(":id/img_check"):
                    child_bounds = _parse_bounds(child)
                    if child_bounds:
                        centers = sum(child_bounds[1::2]) // 2
                        candidates.append((abs(centers - title_center), child_bounds))
            if candidates:
                check_bounds = min(candidates)[1]
        if check_bounds and check_bounds[1] >= 300 and check_bounds[3] <= 2300:
            found[name] = check_bounds
    return found, visible_names


def _select_rows_for_surname(
    device, surname, records, package_name, keyboard_ime=None
):
    initial_count = _selected_count(device, package_name)
    if initial_count != 0:
        return (
            set(),
            {record["name"] for record in records},
            f"页面已有 {initial_count} 人被勾选，已停止以免误选",
        )
    target_names = {record["name"] for record in records}
    _search_box_for_surname(device, surname, keyboard_ime)
    selected = set()
    unchanged_rounds = 0
    previous_visible = set()

    while len(selected) < len(target_names) and unchanged_rounds < 3:
        visible, visible_names = _visible_target_rows(
            device, target_names - selected
        )
        for name, row_bounds in visible.items():
            before = _selected_count(device, package_name)
            left, top, right, bottom = row_bounds
            device.click((left + right) // 2, (top + bottom) // 2)
            time.sleep(0.25)
            after = _selected_count(device, package_name)
            if after == before + 1:
                selected.add(name)

        if len(selected) == len(target_names):
            break

        all_visible = visible_names
        if all_visible == previous_visible:
            unchanged_rounds += 1
        else:
            unchanged_rounds = 0
            previous_visible = all_visible

        try:
            try:
                input_state = device.shell("dumpsys input_method").output
                if "mInputShown=true" in input_state:
                    device.press("back")
                    time.sleep(0.4)
            except Exception:
                pass
            width, height = device.window_size()
            device.swipe(
                width // 2,
                int(height * 0.78),
                width // 2,
                int(height * 0.32),
                0.35,
            )
        except Exception:
            break
        time.sleep(0.35)

    missing = target_names - selected
    if missing and selected:
        try:
            width, height = device.window_size()
            for _ in range(8):
                device.swipe(
                    width // 2,
                    int(height * 0.30),
                    width // 2,
                    int(height * 0.82),
                    0.25,
                )
                time.sleep(0.15)
        except Exception:
            pass

        visible, _ = _visible_target_rows(device, missing)
        for name, row_bounds in visible.items():
            before = _selected_count(device, package_name)
            left, top, right, bottom = row_bounds
            device.click((left + right) // 2, (top + bottom) // 2)
            time.sleep(0.25)
            if _selected_count(device, package_name) == before + 1:
                selected.add(name)
        missing = target_names - selected

    expected = len(selected)
    actual = _selected_count(device, package_name)
    if expected == 0:
        try:
            root, _ = _batch_snapshot(device)
            debug_names = [
                node.attrib.get("text", "").strip()
                for node in root.iter("node")
                if node.attrib.get("text", "").strip()
            ]
            (pathlib.Path.home() / "Desktop").mkdir(exist_ok=True)
            debug_path = pathlib.Path(
                os.environ.get(
                    "DAOMENG_BATCH_DEBUG",
                    str(pathlib.Path.home() / "Desktop" / "batch-debug.txt"),
                )
            )
            debug_path.write_text(
                "search="
                + surname
                + "\n"
                + "\n".join(debug_names[:200]),
                encoding="utf-8",
            )
        except Exception:
            pass
        return set(), target_names, ""
    if actual != expected:
        return (
            set(),
            target_names,
            f"勾选数量不一致，页面显示 {actual} 人，预期 {expected} 人",
        )
    return selected, missing, ""


def _issue_selected_group(device, selected, package_name, selectors):
    award_button = device(
        resourceId=f"{package_name}:id/tv_step",
        text=f"颁发{len(selected)}人",
    )
    if not award_button.exists:
        return set(), f"没有找到“颁发{len(selected)}人”按钮"
    submitted = True
    try:
        device.toast.reset()
    except Exception:
        pass
    award_button.click()
    time.sleep(0.4)
    try:
        confirm = device(**selectors["确定"])
        if confirm.exists:
            confirm.click()
            time.sleep(0.4)
    except Exception:
        pass

    toast = "未获取到Toast"
    try:
        toast = (
            device.toast.get_message(wait_timeout=8, default="未获取到Toast")
            or "未获取到Toast"
        )
    except Exception:
        pass

    count_match = re.search(r"成功\s*(\d+)\s*人", toast)
    if "成功" in toast and (
        not count_match or int(count_match.group(1)) == len(selected)
    ):
        return selected, f"成功批量发分（{toast}）"
    if submitted:
        return selected, f"待核查（提交后状态未知，提示：{toast}）"
    return selected, f"待核查（{toast}）"


def _refresh_award_page(controller, points):
    try:
        controller.device.press("back")
    except Exception:
        pass
    time.sleep(1.0)
    if not controller.navigation(points, "发分界面"):
        raise RuntimeError("重试前无法重新进入发分页面")
    time.sleep(0.5)


def _run_surname_batch():
    import argparse

    parser = argparse.ArgumentParser(description="按姓氏批量发分")
    parser.add_argument("--surname-batch", action="store_true")
    parser.add_argument("--enrollment")
    parser.add_argument("--scores")
    parser.add_argument("--points", type=float, default=2.0)
    parser.add_argument("--one-point", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--select-only", action="store_true")
    parser.add_argument("--select-surname")
    args, _ = parser.parse_known_args(sys.argv[1:])
    executable_name = os.path.basename(sys.executable or sys.argv[0])
    one_point_executable = (
        args.one_point
        or "只发1分" in executable_name
        or "一分专用版" in executable_name
    )

    console.print(
        Panel(
            "1. 选择含“活动名称、姓名、学号”的录取名单。\n"
            "2. 选择本次要发分的人员名单。\n"
            "3. 程序先输出重名和异常清单，再按姓氏分组。\n"
            "4. 逐个勾选后，数量必须与名单完全一致，才会点击“颁发 N 人”。\n"
            "5. 提交结果不明确时立即停止并写入待核查表。\n"
            "6. 启动时可选择“只发1分”或“发2分”模式。",
            title=(
                "按姓氏批量发分 · 只发1分专用版"
                if one_point_executable
                else "按姓氏批量发分"
            ),
            border_style="cyan",
        )
    )
    enrollment_path = pathlib.Path(
        args.enrollment or Prompt.ask("录取名单路径").strip().strip('"').strip("'")
    )
    score_path = pathlib.Path(
        args.scores or Prompt.ask("发分名单路径").strip().strip('"').strip("'")
    )
    points = 1.0 if one_point_executable else args.points
    if not args.enrollment and not args.scores and not one_point_executable:
        mode = Prompt.ask(
            "请选择发放模式：[1] 只发1分  [2] 只发2分  [3] 其他分值",
            choices=["1", "2", "3"],
            default="2",
        )
        if mode == "1":
            points = 1.0
        elif mode == "2":
            points = 2.0
        else:
            try:
                points = float(Prompt.ask("请输入本次分值", default=f"{points:g}"))
            except ValueError:
                raise ValueError("分值必须是数字，例如 1.0")
    elif not args.enrollment and not args.scores:
        console.print("[bold cyan]已选择：只发1分模式[/bold cyan]")
    if not (points >= 0) or points != points or points == float("inf"):
        raise ValueError("分值必须是有限且不小于 0 的数字")
    if not enrollment_path.is_file() or not score_path.is_file():
        raise FileNotFoundError("录取名单或发分名单不存在")

    (
        accepted,
        pending,
        deduplicated,
        groups,
        activity_name,
        duplicate_report,
    ) = _prepare_surname_batches(enrollment_path, score_path)
    output_dir = _batch_output_root()
    pending_path, plan_path, duplicate_path, summary = _write_batch_reports(
        output_dir, accepted, pending, deduplicated, duplicate_report, points
    )
    console.print(
        f"[green]活动：[/green]{activity_name}\n"
        f"[green]可批量发分：[/green]{summary['accepted']} 人\n"
        f"[cyan]姓氏分组：[/cyan]{summary['surname_groups']} 组\n"
        f"[yellow]待核查：[/yellow]{summary['pending']} 人\n"
        f"[magenta]自动去重：[/magenta]{summary['deduplicated']} 条\n"
        f"[magenta]录取名单重名：[/magenta]{summary['duplicate_people']} 人"
        f"（{summary['duplicate_names']} 个姓名）\n"
        f"[green]结果文件夹：[/green]{output_dir}"
    )
    console.print(
        f"[dim]待核查表：{pending_path}\n"
        + (
            f"重名对照表：{duplicate_path}\n"
            if duplicate_path is not None
            else "本次没有两分名单内部重名，未生成重名对照表。\n"
        )
        + f"按姓氏排序名单：{plan_path}[/dim]"
    )
    try:
        os.startfile(str(output_dir))
    except Exception:
        pass

    if args.dry_run:
        _write_batch_results(
            output_dir,
            accepted,
            pending,
            points,
            ["仅完成名单分析和导出，没有连接手机，也没有发分。"],
        )
        console.print("[cyan]预览模式完成，未连接手机，未发分。[/cyan]")
        return 0

    if not args.select_only:
        confirmation = Prompt.ask(
            "确认开始批量发分？输入 START 继续（直接回车取消）",
            default="",
        ).strip()
        if confirmation != "START":
            pending.extend(accepted)
            for record in accepted:
                record["reason"] = "用户取消执行"
                record["status"] = "未执行"
            accepted.clear()
            _write_batch_results(
                output_dir,
                accepted,
                pending,
                points,
                ["用户取消，未连接手机。"],
            )
            return 0

    device_module = importlib.import_module("daomeng.device_controller")
    selector_module = importlib.import_module("daomeng.selectors")
    config_module = importlib.import_module("daomeng.config")
    logging_module = importlib.import_module("daomeng.logging_utils")
    config = config_module.AutomationConfig(
        device=None,
        pair=None,
        additional_list=None,
        console_log_level="INFO",
        file_log_level="INFO",
        file_dir={0.0: [enrollment_path], points: [score_path]},
        output_dir=output_dir,
        init_device=False,
        accounts=None,
        activity_name=activity_name,
        yiban_mapping=None,
        click_other_tab=False,
    )
    logger = logging_module.get_logger(
        "INFO", output_dir / "automation.log", "INFO"
    )
    controller = device_module.DeviceController(config, logger)
    run_notes = []
    try:
        controller.connect()
        if not controller.navigation(points, "发分界面"):
            raise RuntimeError("没有到达发分页面")
        try:
            original_ime = (
                controller.device.shell(
                    "settings get secure default_input_method"
                ).output.strip()
            )
        except Exception:
            original_ime = None
        with controller.fast_input():
            if args.select_only:
                if args.select_surname:
                    surname = args.select_surname.strip()
                    if surname not in groups:
                        raise ValueError(f"没有可勾选的姓氏组：{surname}")
                    records = groups[surname]
                else:
                    surname, records = next(iter(groups.items()))
                selected_names, missing_names, select_error = _select_rows_for_surname(
                    controller.device,
                    surname,
                    records,
                    selector_module.DMKJ_PACKAGE_NAME,
                    original_ime,
                )
                if select_error:
                    raise RuntimeError(select_error)
                console.print(
                    f"[green]选择验证完成：姓氏 {surname} 已勾选 "
                    f"{len(selected_names)} 人，二次查找仍未找到 "
                    f"{len(missing_names)} 人；未点击“颁发 N 人”。[/green]"
                )
                return 0
            batch_completed = True
            for surname, records in groups.items():
                retry_count = 0
                final_result = ""
                while True:
                    if (
                        _selected_count(
                            controller.device,
                            selector_module.DMKJ_PACKAGE_NAME,
                        )
                        != 0
                    ):
                        _refresh_award_page(controller, points)
                    selected_names, missing_names, select_error = _select_rows_for_surname(
                        controller.device,
                        surname,
                        records,
                        selector_module.DMKJ_PACKAGE_NAME,
                        original_ime,
                    )
                    if select_error:
                        final_result = select_error
                        selected_names = set()
                        break
                    if not selected_names:
                        final_result = "姓氏搜索没有找到可勾选人员"
                        break

                    selected_names, final_result = _issue_selected_group(
                        controller.device,
                        selected_names,
                        selector_module.DMKJ_PACKAGE_NAME,
                        selector_module.SELECTORS,
                    )
                    if (
                        "请求数据失败" in final_result
                        and retry_count < 3
                    ):
                        retry_count += 1
                        note = (
                            f"姓氏 {surname} 请求失败，"
                            f"等待20秒后进行第 {retry_count}/3 次重试"
                        )
                        run_notes.append(note)
                        logger.warning(note)
                        console.print(f"[yellow]{note}[/yellow]")
                        _write_batch_results(
                            output_dir,
                            accepted,
                            pending,
                            points,
                            run_notes,
                        )
                        time.sleep(20)
                        _refresh_award_page(controller, points)
                        continue
                    break

                if "成功批量发分" in final_result:
                    for record in records:
                        if record["name"] in selected_names:
                            record["status"] = f"成功批量发分{points:g}分"
                        else:
                            record["status"] = (
                                "待核查（顶部二次查找仍未见）"
                            )
                elif "姓氏搜索没有找到可勾选人员" in final_result:
                    for record in records:
                        record["status"] = "待核查（姓氏搜索未找到）"
                else:
                    for record in records:
                        record["status"] = final_result
                    if "请求数据失败" in final_result:
                        console.print(
                            "[yellow]请求连续失败 3 次，30 秒后关闭程序并锁屏。"
                            "请先人工核对本姓氏人员是否实际获得学分。[/yellow]"
                        )
                        time.sleep(30)
                    if "状态未知" in final_result:
                        batch_completed = False
                        run_notes.append(
                            f"姓氏 {surname} 停止：{final_result}"
                        )
                        _write_batch_results(
                            output_dir,
                            accepted,
                            pending,
                            points,
                            run_notes,
                        )
                        break

                run_notes.append(f"姓氏 {surname}：{final_result}")
                _write_batch_results(
                    output_dir,
                    accepted,
                    pending,
                    points,
                    run_notes,
                )
            if batch_completed:
                console.print()
                console.print(
                    Panel(
                        "[bold green]本次按姓氏批量发分流程已结束。[/bold green]\n\n"
                        "感谢使用本工具。\n"
                        "[bold cyan]鸣谢usc易班工作站技术部and行政部[/bold cyan]",
                        title="运行结束",
                        border_style="green",
                    )
                )
    finally:
        try:
            controller.cleanup()
        except Exception:
            pass
        _write_batch_results(
            output_dir, accepted, pending, points, run_notes
        )
        try:
            os.startfile(str(output_dir))
        except Exception:
            pass
    return 0


def _install_navigation_fix() -> None:
    device_module = importlib.import_module("daomeng.device_controller")
    selector_module = importlib.import_module("daomeng.selectors")
    controller_class = device_module.DeviceController
    selectors = selector_module.SELECTORS
    package_name = selector_module.DMKJ_PACKAGE_NAME

    def current_package_from_system(device):
        try:
            response = device.shell("dumpsys window")
            output = getattr(response, "output", str(response))
            matches = re.findall(
                r"mCurrentFocus=Window\{[^}]*?\s+([A-Za-z0-9._]+)/",
                output or "",
            )
            if package_name in matches:
                return package_name
            if matches:
                return matches[-1]
        except Exception:
            pass
        try:
            root = ET.fromstring(device.dump_hierarchy())
            return str(root.attrib.get("package", ""))
        except Exception:
            return ""

    def wait_for_dream_app(controller, timeout=25):
        deadline = time.monotonic() + timeout
        attempts = 0
        last_package = ""
        try:
            controller.device.shell(
                "input keyevent KEYCODE_WAKEUP; wm dismiss-keyguard"
            )
        except Exception:
            pass
        while time.monotonic() < deadline:
            try:
                current = controller.device.app_current() or {}
                last_package = str(current.get("package", ""))
            except Exception as exc:
                last_package = f"读取失败: {exc}"

            if last_package == package_name:
                return

            # Never use stop=True while the app is still loading. Repeatedly
            # stopping it was the reason the previous version stayed on N0.
            if attempts < 2 and time.monotonic() + 5 < deadline:
                attempts += 1
                try:
                    controller.device.app_start(package_name)
                except Exception:
                    pass
                time.sleep(1.0)
                continue
            time.sleep(0.5)

        raise RuntimeError(
            "没有进入目标应用。当前前台包名是 "
            f"{last_package or '未知'}，期望 {package_name}。"
            "请确认应用已安装、手机解锁并允许 USB 调试。"
        )

    def wait_loading_gone(controller, timeout=15):
        try:
            controller.device(**selectors["加载"]).wait_gone(timeout=timeout)
        except Exception:
            pass

    def tap_element(controller, element):
        bounds = str(element.info.get("bounds", ""))
        match = re.fullmatch(r"\[(\d+),(\d+)\]\[(\d+),(\d+)\]", bounds)
        if not match:
            element.click()
            return
        left, top, right, bottom = map(int, match.groups())
        controller.device.click((left + right) // 2, (top + bottom) // 2)

    def bounds(node):
        match = re.fullmatch(
            r"\[(\d+),(\d+)\]\[(\d+),(\d+)\]", node.attrib.get("bounds", "")
        )
        if not match:
            return None
        return tuple(map(int, match.groups()))

    def click_manage_for_activity(controller, activity_name):
        d = controller.device
        try:
            root = ET.fromstring(d.dump_hierarchy())
        except Exception:
            return False

        title_nodes = [
            node
            for node in root.iter("node")
            if node.attrib.get("resource-id", "").endswith(":id/tv_activity_name")
            and node.attrib.get("text", "").strip() == activity_name.strip()
        ]
        manage_nodes = [
            node
            for node in root.iter("node")
            if node.attrib.get("resource-id", "").endswith(":id/tv_tag_bb")
            and node.attrib.get("text", "").strip() == "管理活动"
        ]
        for title in title_nodes:
            title_bounds = bounds(title)
            if not title_bounds:
                continue
            title_bottom = title_bounds[3]
            below = []
            for manage in manage_nodes:
                manage_bounds = bounds(manage)
                if manage_bounds and manage_bounds[1] >= title_bottom:
                    below.append((manage_bounds[1] - title_bottom, manage_bounds))
            if not below:
                continue
            _, nearest = min(below)
            left, top, right, bottom = nearest
            d.click((left + right) // 2, (top + bottom) // 2)
            return True
        return False

    def select_filter(controller, status):
        d = controller.device
        condition = d(resourceId=f"{package_name}:id/radio_condition")
        if not condition.exists:
            return False
        if condition.get_text() == status:
            return True
        tap_element(controller, condition)
        time.sleep(0.25)
        option = d(resourceId=f"{package_name}:id/title", text=status)
        if not option.exists:
            return False
        # Filter menu entries are TextViews rather than clickable buttons.
        # UiObject.click() silently does nothing on this page, so tap the
        # center of their reported bounds explicitly.
        tap_element(controller, option)
        close_button = d(resourceId=f"{package_name}:id/imgbtn_close")
        deadline = time.monotonic() + 4
        while close_button.exists and time.monotonic() < deadline:
            time.sleep(0.2)
        if close_button.exists:
            d.press("back")
        wait_loading_gone(controller)
        return True

    def find_activity(controller, activity_name):
        d = controller.device
        manage_tab = d(**selectors["活动管理tab"])
        if manage_tab.exists:
            try:
                if not manage_tab.info.get("checked", False):
                    manage_tab.click()
                    wait_loading_gone(controller)
            except Exception:
                manage_tab.click()
                wait_loading_gone(controller)

        search_box = d(**selectors["活动搜索框"])
        deadline = time.monotonic() + 10
        while not search_box.exists and time.monotonic() < deadline:
            time.sleep(0.3)
        if not search_box.exists:
            raise RuntimeError(
                "已经点击“我的活动 → 管理”，但页面没有出现活动搜索框。"
                "请确认当前账号具有活动管理权限。"
            )

        def search_once():
            search_box.set_text(activity_name)
            d.press("enter")
            wait_loading_gone(controller)
            return click_manage_for_activity(controller, activity_name)

        statuses = [
            "已结束",
            "已完结",
            "进行中",
            "等待中",
            "报名中",
            "规划中",
            "待审核",
            "审核驳回",
        ]
        checked = []
        if click_manage_for_activity(controller, activity_name):
            console.print("[green]活动管理页已找到目标活动，使用当前筛选[/green]")
            return True

        if search_once():
            console.print("[green]活动管理页已找到活动，使用当前筛选[/green]")
            return True

        for status in statuses:
            if select_filter(controller, status):
                checked.append(status)

            if search_once():
                console.print(
                    f"[green]活动管理页已找到活动，筛选状态：{status}[/green]"
                )
                return True

        raise RuntimeError(
            "已经进入“我的活动 → 管理”并完成搜索，但该账号下没有匹配的活动。"
            f"\n活动名称：{activity_name}"
            f"\n已检查状态：{', '.join(checked) or '当前筛选'}"
            "\n请核对 Excel 里的活动名称是否与手机端完全一致，"
            "并确认当前登录账号有该活动的管理权限。"
        )

    def fixed_navigation(self, score, target="发分界面"):
        if not self.device:
            raise RuntimeError("设备未连接，无法执行导航")

        # OnePlus/Android 16 can expose a floating game panel as the window
        # focus even while DreamSpace is visibly active. The package embedded
        # in the current UI hierarchy is the reliable signal.
        try:
            if not getattr(self.device, "_daomeng_ui_package_fix", False):
                self.device.app_current = lambda: {
                    "package": current_package_from_system(self.device)
                }
                self.device._daomeng_ui_package_fix = True
        except Exception:
            pass

        goal = self._target_state(score, target)
        back_count = 0
        score_entry_attempts = 0
        wait_for_dream_app(self)
        self._logger.info("等待到梦空间进入前台，目标界面: %s", goal)

        while back_count <= 60:
            time.sleep(0.3)
            current = self._get_current()
            if current == goal:
                self._logger.info("成功导航到目标界面: %s", goal)
                return True

            if current == "N0":
                wait_for_dream_app(self)
                continue

            if target != "发分界面":
                if current == target:
                    return True
                self.device.press("back")
                back_count += 1
                continue

            if current == "登录页面":
                if not self._config.accounts:
                    self._pause_live_for_prompt()
                    try:
                        from rich.prompt import Prompt

                        Prompt.ask("检测到需要登录，请登录后按回车继续")
                    finally:
                        self._resume_live()
                    continue

                accounts = self._config.accounts
                username = list(accounts.keys())[0]
                password = list(accounts.values())[0]
                self.device(**selectors["登录账号"]).set_text(username)
                self.device(**selectors["登录密码"]).set_text(password)
                self.device.hide_keyboard()
                self.device(**selectors["同意协议"]).click()
                self.device(**selectors["登录按钮"]).click()
                wait_loading_gone(self)
                continue

            if current == "首页":
                self.device(**selectors["我的"]).click()
                continue

            if current == "我的":
                self.device(**selectors["我的活动"]).click()
                wait_loading_gone(self)
                continue

            if current == "活动管理":
                find_activity(self, self._config.activity_name)
                wait_loading_gone(self)
                continue

            if current == "管理列表":
                self.device(**selectors["学分管理"]).click()
                wait_loading_gone(self)
                continue

            if current == "学分管理":
                score_entry_attempts += 1
                if score_entry_attempts > 2:
                    raise RuntimeError(
                        "已点击进入发分页两次，但页面状态仍未切换。"
                        "程序已停止继续点击，避免反复进入；请手动检查当前页面。"
                    )
                self.score = score
                credit = self.device(
                    **selectors["积分"], text=f"积分 {score:.1f}"
                )
                tap_element(self, credit)
                time.sleep(0.5)
                self.device(**selectors["颁发"]).click()
                wait_loading_gone(self)
                current = self._get_current()
                if current == f"{score}分发放界面":
                    self.score = score
                    return True
                continue

            if current.startswith("分发放界面"):
                if self.score == score:
                    return True
                self.device.press("back")
                back_count += 1
                continue

            if current == "N1":
                self.device.press("back")
                back_count += 1
                continue

            self.device.press("back")
            back_count += 1

        raise RuntimeError(
            f"已进入活动管理流程，但未能到达 {goal}。"
            "程序没有执行任何颁发操作，请手动检查当前页面。"
        )

    controller_class.navigation = fixed_navigation
    original_get_current = controller_class._get_current_nolog

    def fixed_get_current(self):
        # The student award page contains both the "颁发" control and the
        # "未获得任何学分" tab. Checking "颁发" first classifies that page as
        # 学分管理 and causes navigation to click into it repeatedly.
        try:
            if self.device and self.device(**selectors["未获得任何学分"]).exists:
                return f"{self.score}分发放界面"
        except Exception:
            pass
        return original_get_current(self)

    controller_class._get_current_nolog = fixed_get_current

    def fixed_award_points(self, student):
        if not self.device:
            self._logger.error("设备未连接，无法颁发学分")
            raise RuntimeError("设备未连接")

        attempt = 0
        while True:
            d = self.device
            submitted = False
            try:
                self.score = self._score()
                if not self.score or self.score != student.points:
                    if not self.navigation(student.points, "发分界面"):
                        raise RuntimeError("导航到发分界面失败")
                    d = self.device

                if self._config.click_other_tab:
                    other_tab = d(**selectors["已获得其他学分"])
                    if other_tab.exists:
                        other_tab.click()
                        wait_loading_gone(self)

                row = self._search_student(student)
                if row is None:
                    return student

                d(**selectors["颁发1人"]).click()

                # From this point onward the server may have accepted the award.
                # Never retry automatically because the UI may fail before it can
                # display a loading state or a success Toast.
                submitted = True
                d(**selectors["确定"]).click()
                wait_loading_gone(self)
                try:
                    d(**selectors["发分加载"]).wait_gone(timeout=20)
                except Exception:
                    pass
                toast = (
                    d.toast.get_message(
                        wait_timeout=7, default="未获取到Toast"
                    )
                    or "未获取到Toast"
                )
                if "成功1人" in toast:
                    student.status = f"成功颁发{student.points:.1f}分"
                    self._logger.info(
                        "确认颁发成功: %s (%s), %.1f 分",
                        student.name,
                        student.student_id,
                        student.points,
                    )
                else:
                    student.status = f"待核对（提交后提示:{toast}）"
                    self._logger.warning(
                        "提交后未能确认发放结果: %s (%s), Toast=%s",
                        student.name,
                        student.student_id,
                        toast,
                    )
                return student
            except Exception as exc:
                if submitted:
                    student.status = "待核对（提交后状态未知）"
                    self._logger.exception(
                        "提交发分后发生异常，不再自动重试: %s (%s)",
                        student.name,
                        student.student_id,
                    )
                    return student

                if self._is_device_connection_error(exc):
                    self._logger.error("发分前设备连接失败: %s", exc)
                    try:
                        self.connect()
                        d = self.device
                        self._logger.info("设备自动重连成功，将继续发分")
                    except Exception as reconnect_error:
                        self._logger.error("设备自动重连失败: %s", reconnect_error)
                        if not self._reconnect_with_prompt():
                            student.status = "待核对（发分前用户退出）"
                            raise RuntimeError("用户选择退出") from exc
                        d = self.device

                if attempt > 2:
                    student.status = f"发分失败（{exc}）"
                    self._logger.error(
                        "为 %s (%s) 发分失败，已尝试 %d 次: %s",
                        student.name,
                        student.student_id,
                        attempt,
                        exc,
                    )
                    return student

                self._logger.warning(
                    "发分前异常，准备重试 %s (%s): %s",
                    student.name,
                    student.student_id,
                    exc,
                )
                attempt += 1
                time.sleep(0.5)

    controller_class.award_points = fixed_award_points

    def fixed_search_student(self, student):
        d = self.device
        search_box = d(**selectors["人员搜索框"])

        def collect_exact_matches():
            try:
                root = ET.fromstring(d.dump_hierarchy())
            except Exception:
                return []
            return [
                node
                for node in root.iter("node")
                if node.attrib.get("resource-id", "").endswith(":id/title")
                and node.attrib.get("text", "").strip() == student.name.strip()
                and bounds(node)
            ]

        actual_query = ""
        for search_attempt in range(3):
            deadline = time.monotonic() + 10
            while not search_box.exists and time.monotonic() < deadline:
                time.sleep(0.3)
            if not search_box.exists:
                self._logger.warning(
                    "人员搜索框暂时未出现，重试 %d/3: %s (%s)",
                    search_attempt + 1,
                    student.name,
                    student.student_id,
                )
                time.sleep(0.5)
                continue

            _ensure_system_keyboard(d)
            time.sleep(0.5)
            if not _show_system_keyboard(d, search_box):
                raise RuntimeError("系统输入法键盘没有弹出")
            search_box.set_text(student.student_id)
            time.sleep(0.35)
            try:
                actual_query = search_box.get_text().strip()
            except Exception:
                actual_query = ""
            if student.student_id not in actual_query:
                self._logger.warning(
                    "人员学号写入未确认，重试 %d/3: %s (%s)，当前='%s'",
                    search_attempt + 1,
                    student.name,
                    student.student_id,
                    actual_query,
                )
                continue

            _click_keyboard_search(d)
            wait_loading_gone(self)

            deadline = time.monotonic() + 6
            matched_nodes = []
            while time.monotonic() < deadline:
                matched_nodes = collect_exact_matches()
                if matched_nodes:
                    break
                time.sleep(0.3)
            if matched_nodes:
                if len(matched_nodes) == 1:
                    left, top, right, bottom = bounds(matched_nodes[0])
                    d.click((left + right) // 2, (top + bottom) // 2)
                    time.sleep(0.4)
                    return True
                student.status = f"待核对（同名结果数量={len(matched_nodes)}）"
                return None
            if d(**selectors["无结果"]).exists:
                student.status = "待核查（学号搜索无结果，已跳过）"
                return None
            student.status = "待核查（搜索到结果但姓名不符，已跳过）"
            return None

        student.status = "待核对（人员搜索框未确认）"
        self._logger.warning(
            "人员搜索框连续三次未确认，停止操作: %s (%s)，当前='%s'",
            student.name,
            student.student_id,
            actual_query,
        )
        return None

    controller_class._search_student = fixed_search_student

    @contextlib.contextmanager
    def no_fast_input(self):
        yield

    controller_class.fast_input = no_fast_input
    if os.environ.get("DAOMENG_HOLD_AWARD") == "1":
        controller_class.cleanup = lambda self: None


_install_navigation_fix()


def _pause_before_close(code: int) -> None:
    if not getattr(sys, "frozen", False):
        return
    if "--no-pause" in sys.argv:
        return
    if code == 0 and not interactive.wizard_used:
        return
    stream = sys.stdout
    if stream is None or not stream.isatty():
        return
    try:
        console.print("\n[dim]按回车键关闭窗口...[/dim]")
        input()
    except Exception:
        pass


def main() -> int:
    executable_name = os.path.basename(sys.executable or sys.argv[0])
    if "--surname-batch" in sys.argv or "按姓氏" in executable_name:
        batch_code = 0
        try:
            batch_code = _run_surname_batch()
        except Exception:
            console.print_exception()
            batch_code = 1
        finally:
            _pause_before_close(batch_code if batch_code != 0 else 1)
        return batch_code

    code = 0
    try:
        code = run_cli()
    except KeyboardInterrupt:
        console.print("\n[bold red]程序已被用户中断。[/bold red]")
        code = 1
    except Exception:
        console.print_exception()
        code = 1
    finally:
        _pause_before_close(code)
    return code


if __name__ == "__main__":
    sys.exit(main())
