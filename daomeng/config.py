from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from rich.progress import (
    Progress,
    MofNCompleteColumn,
    SpinnerColumn,
    BarColumn,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn)
        
        
@dataclass
class AutomationConfig:
    """运行时配置。"""

    device: Optional[str]
    pair: Optional[list[str]]
    additional_list: Optional[Path]
    console_log_level: str
    file_log_level: str
    file_dir: dict[float, Path]
    output_dir: Path
    init_device: bool
    accounts: Optional[dict[str, str]]
    activity_name: str
    yiban_mapping: Optional[Path]
    click_other_tab: bool = False


@dataclass
class Student:
    """学生信息。"""

    student_id: str
    name: str
    points: float
    status: str
    
    
def progress_():
    return Progress(
    SpinnerColumn(),
    TextColumn("[progress.description]{task.description}"),
    BarColumn(bar_width=None),
    MofNCompleteColumn(),
    TimeElapsedColumn(),
    TimeRemainingColumn(),
    expand=True,
)