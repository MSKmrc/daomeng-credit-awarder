import logging
import time
from rich.live import Live
from rich.console import Group
from rich.panel import Panel
from rich.text import Text

from .config import AutomationConfig,progress_
from .file_manager import FileManager
from .device_controller import DeviceController


class AutomationApp:
    def __init__(self, config: AutomationConfig, logger: logging.Logger) -> None:
        self.config = config
        self.logger = logger
        self.device_controller = DeviceController(config, logger)

    def run(self) -> int:
        try:
            self.logger.info("="*60)
            self.logger.info("开始进行设备连接")
            self.device_controller.connect()
            self.logger.info("设备连接完成，设备已就绪")
            self.logger.info("="*60)
            
            with FileManager(self.config, self.logger) as file_manager:
                # 检查数据库表是否存在，而不仅仅是文件
                if file_manager.check_database_initialized():
                    self.logger.info("检测到已有学生数据库，跳过Excel读取步骤")
                else:
                    self.logger.info("开始读取Excel文件并生成学生数据库")
                    file_manager.save_students_to_db()
                    self.logger.info("学生数据库生成完成")
                
                points_list = file_manager.get_all_points_list()
                self.logger.info(f"共读取到 {len(points_list)} 个不同分值: {points_list}")
                
                total_students = file_manager.get_unprocessed_students_count()
                self.logger.info(f"待处理学生总数: {total_students}")
                self.logger.info("="*60)
                
                progress = progress_()
                status_text_1 = Text("准备开始...", style="italic dim")
                with Live(
                    Panel(Group(progress, status_text_1)),
                    refresh_per_second=10,
                    transient=False,
                ) as live:
                    # 将 Live 对象传递给 DeviceController，以便在需要用户交互时暂停显示
                    self.device_controller.set_live(live)
                    
                    with self.device_controller.fast_input():
                        self.logger.info("开始运行自动化任务")
                        self.device_controller.device.unlock()
                        progress_task_1 = progress.add_task("总进度", total=total_students)
                        done_count = 0
                        overall_start = time.time()

                        def _eta_text() -> str:
                            """根据已完成学生的平均耗时估算剩余时间。"""
                            if done_count == 0:
                                return ""
                            avg = (time.time() - overall_start) / done_count
                            left = max(total_students - done_count, 0)
                            secs = int(avg * left)
                            h, rem = divmod(secs, 3600)
                            m, s = divmod(rem, 60)
                            eta = f"{h}小时{m:02d}分" if h else (f"{m}分{s:02d}秒" if m else f"{s}秒")
                            return f"｜平均 {avg:.1f} 秒/人｜剩余 {left} 人，预计还需 {eta}"

                        for point in points_list:
                            self.logger.info("="*60)
                            self.logger.info(f"开始处理分值: {point:.1f}")
                            self.device_controller.navigation(point)
                            student_list = file_manager.get_unprocessed_students_by_points(point)
                            self.logger.info(f"该分值待处理学生数: {len(student_list)}")
                            progress_task_2 = progress.add_task(f"颁发 {point:.1f} 分", total=len(student_list))
                            for student in student_list:
                                status_text_1.plain = f"当前处理学生: {student.student_id}，姓名: {student.name}，学分: {student.points:.1f}"
                                started_at = time.time()
                                student = self.device_controller.award_points(student)
                                file_manager.update_students_status([student])
                                cost = time.time() - started_at
                                done_count += 1
                                self.logger.debug(
                                    f"学生 {student.student_id} 耗时 {cost:.2f} 秒，状态: {student.status}"
                                )
                                status_text_1.plain = (
                                    f"刚处理: {student.student_id} {student.name} "
                                    f"（{student.points:.1f} 分，{cost:.1f} 秒）{_eta_text()}"
                                )
                                progress.update(progress_task_1, advance=1)
                                progress.update(progress_task_2, advance=1)
                            progress.remove_task(progress_task_2)
                            self.logger.info(f"分值 {point:.1f} 处理完成")
                        status_text_1.plain = "全部完成 ✅"
                    
                    # 清除 DeviceController 中的 Live 引用
                    self.device_controller.set_live(None)
                    
                self.logger.info("="*60)
                self.logger.info("自动化任务成功完成")
                self.logger.info("="*60)
                file_manager.export_db_to_excel()
            file_manager.cleanup()
            
            return 0
        except Exception as e:
            self.logger.error("="*60)
            self.logger.error(f"自动化任务失败: {e}")
            self.logger.exception("详细错误信息:")
            self.logger.error("="*60)
            return 1
        
        finally:
            self.device_controller.cleanup()
            