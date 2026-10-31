# 到梦空间自动化脚本 - AI 代码助手指南

## 项目概述

这是一个基于 Android UI 自动化的学分颁发系统，通过 ADB 和 uiautomator2 控制移动设备，自动为学生批量颁发"到梦空间"APP 的活动学分。

**核心价值链**: Excel 名单 → 学号验证 → SQLite 持久化 → UI 自动导航 → 学分颁发 → 结果跟踪

## 架构设计
## 到梦空间自动化脚本 — AI 代码助手快速指南

目标：让 AI 代理能快速理解本仓库的架构、数据流、关键约定和常见修改点，以便安全、可重复地编辑/调试代码。

主要入口与文件（优先级高）
- `DM.py`：命令行入口
- `daomeng/cli.py`：参数解析与任务恢复（`-r`、`-p`、`-u`、`-yb`）
- `daomeng/automation_app.py`：任务编排、进度与错误捕获（主要协调器）
- `daomeng/file_manager.py`：Excel I/O、学号提取、SQLite 持久化（查看 `_extract_student_id()`）
- `daomeng/device_controller.py`：ADB/uiautomator2 会话、`navigation()` 状态机、`award_points()` 核心逻辑
- `daomeng/selectors.py`：所有 UI 选择器集中在 `SELECTORS` 字典，任何 UI 更改请先在此处修改

快速事实与约定（写代码时务必遵守）
- 学号提取使用正则示例：`r'(?<!\d)\d{11}(?!\d)'`（见 `file_manager._extract_student_id`）
- 日志请使用实例日志器 `self._logger`（不要用 print），用户交互用 `rich.Console()`
- ADB 路径由环境变量 `ADBUTILS_ADB_PATH` 指向 `platform-tools/adb.exe`
- 运行会在 `<活动名>-<时间戳>/` 生成 `config.json` 与 `students.db`；恢复模式 `-r <dir>` 会读取该目录跳过 Excel
- 完成后会删除 `.db`/`.db-wal`/`.db-shm`/`config.json`，仅保留导出的 Excel

常见修改点（示例）
- 添加 UI 页面：编辑 `daomeng/selectors.py` → 更新 `device_controller._get_current()` → 在 `navigation()` 添加分支
- 修改学号规则：编辑 `daomeng/file_manager.py::_extract_student_id()` 并保留 11 位判定逻辑
- 调整颁发重试：查看 `device_controller.award_points()` 中的重试计数和异常处理

调试建议（设备/连接）
- 确认 `platform-tools/adb.exe` 存在并且 `ADBUTILS_ADB_PATH` 指向它
- 无线调试：先用 `-p IP:PORT CODE` 配对，再用 `-d IP:PORT` 连接
- 检查 `automation.log`（仓库根或运行目录）获取 ADB/uiautomator2 异常上下文

数据/输入约定（Excel/CSV）
- 录取名单（score=0）必须含 `学号`、`姓名`
- 发分名单必须含 `姓名`；若无 `学号`，可通过 `-yb` 指定易班映射 CSV（需含 `姓名,学号,易班ID`）

安全与边界条件
- 不支持并发多设备（全局 `self.score` 假设单设备）
- 若 UI 选择器失效，应先用 `uiautomatorviewer` 确认新 selector，再提交改动

如果你需要我把这些要点进一步浓缩为 PR 模板、CODEOWNERS 或添加编辑示例（小补丁）请告诉我要侧重的部分。
