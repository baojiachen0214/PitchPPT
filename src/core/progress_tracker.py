"""
精细化进度跟踪系统
提供平滑、详细的进度更新
"""

import threading
from typing import Callable, Optional
from enum import Enum, auto


class ConversionStage(Enum):
    """转换阶段枚举"""
    INITIALIZING = auto()      # 初始化阶段 (0-5%)
    OPENING_FILE = auto()      # 打开文件 (5-10%)
    ANALYZING = auto()         # 分析PPT (10-15%)
    EXPORTING_IMAGES = auto()  # 导出图片 (15-50%)
    PROCESSING_SLIDES = auto() # 处理幻灯片 (50-85%)
    SAVING = auto()            # 保存文件 (85-95%)
    CLEANING = auto()          # 清理资源 (95-100%)


class StageProgress:
    """阶段进度管理"""
    
    def __init__(self, start_percent: float, end_percent: float, description: str):
        self.start_percent = start_percent
        self.end_percent = end_percent
        self.range = end_percent - start_percent
        self.description = description
        self.current_step = 0
        self.total_steps = 1
    
    def update(self, step: int = None, total_steps: int = None, sub_description: str = None) -> tuple:
        """
        更新阶段内进度
        
        Returns:
            tuple: (总进度百分比, 描述文本)
        """
        if step is not None:
            self.current_step = step
        if total_steps is not None:
            self.total_steps = max(1, total_steps)
        
        # 计算阶段内进度 (0-1)
        stage_progress = min(1.0, self.current_step / self.total_steps)
        
        # 计算总进度
        total_progress = self.start_percent + (self.range * stage_progress)
        
        # 构建描述
        desc = self.description
        if sub_description:
            desc = f"{desc} - {sub_description}"
        if self.total_steps > 1:
            desc = f"{desc} ({self.current_step}/{self.total_steps})"
        
        return total_progress, desc


class ProgressTracker:
    """
    精细化进度跟踪器
    管理转换过程中的详细进度更新
    """
    
    def __init__(self, callback: Callable[[float, str], None] = None):
        self.callback = callback
        self._lock = threading.Lock()
        self._current_stage = None
        
        # 定义各阶段及其进度范围
        self._stages = {
            ConversionStage.INITIALIZING: StageProgress(0.0, 0.05, "初始化PowerPoint"),
            ConversionStage.OPENING_FILE: StageProgress(0.05, 0.10, "打开PPT文件"),
            ConversionStage.ANALYZING: StageProgress(0.10, 0.15, "分析PPT结构"),
            ConversionStage.EXPORTING_IMAGES: StageProgress(0.15, 0.50, "导出幻灯片图片"),
            ConversionStage.PROCESSING_SLIDES: StageProgress(0.50, 0.85, "处理幻灯片背景"),
            ConversionStage.SAVING: StageProgress(0.85, 0.95, "保存输出文件"),
            ConversionStage.CLEANING: StageProgress(0.95, 1.00, "清理临时资源"),
        }
    
    def start_stage(self, stage: ConversionStage, total_steps: int = 1):
        """开始一个新阶段"""
        with self._lock:
            self._current_stage = stage
            stage_progress = self._stages[stage]
            stage_progress.current_step = 0
            stage_progress.total_steps = total_steps
        
        # 立即报告阶段开始 - 显示该阶段的起始进度，而不是0
        start_progress = stage_progress.start_percent
        self._report_progress(start_progress, f"开始{stage_progress.description}")
    
    def update_stage(self, step: int, sub_description: str = None):
        """更新当前阶段的进度"""
        with self._lock:
            if self._current_stage is None:
                return
            
            stage_progress = self._stages[self._current_stage]
            progress, description = stage_progress.update(step, sub_description=sub_description)
        
        self._report_progress(progress, description)
    
    def step(self, sub_description: str = None):
        """前进一步"""
        with self._lock:
            if self._current_stage is None:
                return
            
            stage_progress = self._stages[self._current_stage]
            stage_progress.current_step += 1
            progress, description = stage_progress.update(sub_description=sub_description)
        
        self._report_progress(progress, description)
    
    def finish_stage(self, sub_description: str = None):
        """完成当前阶段"""
        with self._lock:
            if self._current_stage is None:
                return
            
            stage_progress = self._stages[self._current_stage]
            stage_progress.current_step = stage_progress.total_steps
            progress, description = stage_progress.update(sub_description=sub_description)
            
        self._report_progress(progress, description)
    
    def _report_progress(self, progress: float, description: str):
        """报告进度"""
        if self.callback:
            self.callback(progress, description)
    
    def complete(self, success: bool = True):
        """标记完成"""
        desc = "转换完成" if success else "转换失败"
        self._report_progress(1.0 if success else 0.0, desc)
    
    def get_current_progress(self) -> tuple:
        """获取当前进度"""
        with self._lock:
            if self._current_stage is None:
                return 0.0, "准备中..."
            
            stage_progress = self._stages[self._current_stage]
            return stage_progress.update()
