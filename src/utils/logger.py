import logging
import os
from datetime import datetime
from logging.handlers import RotatingFileHandler
from typing import Optional
import time

class PerformanceLogger:
    """
    性能日志记录器
    用于记录操作耗时和性能指标
    """
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self._start_times = {}
    
    def start_timer(self, operation: str):
        """
        开始计时
        
        Args:
            operation: 操作名称
        """
        self._start_times[operation] = time.time()
        self.logger.debug(f"开始计时: {operation}")
    
    def end_timer(self, operation: str):
        """
        结束计时并记录
        
        Args:
            operation: 操作名称
        """
        if operation in self._start_times:
            elapsed = time.time() - self._start_times[operation]
            self.logger.info(f"操作完成: {operation}, 耗时: {elapsed:.3f}秒")
            del self._start_times[operation]
            return elapsed
        return None

class Logger:
    """
    增强的日志系统
    支持日志轮转、多级别日志、性能日志
    """
    
    _instance = None
    _logger = None
    
    def __new__(cls, log_dir: str = "logs", max_bytes: int = 10 * 1024 * 1024, backup_count: int = 5):
        """
        单例模式确保全局只有一个Logger实例
        """
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialize(log_dir, max_bytes, backup_count)
        return cls._instance
    
    def _initialize(self, log_dir: str, max_bytes: int, backup_count: int):
        """
        初始化日志系统
        """
        self.log_dir = log_dir
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # 创建日志文件名（包含日期）
        timestamp = datetime.now().strftime("%Y%m%d")
        log_file = os.path.join(log_dir, f"pitchppt_{timestamp}.log")
        
        # 配置日志格式
        detailed_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(funcName)s() - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        simple_formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%H:%M:%S'
        )
        
        # 获取根记录器并设置级别
        self.logger = logging.getLogger('PitchPPT')
        self.logger.setLevel(logging.DEBUG)
        
        # 清除已有的处理器
        self.logger.handlers.clear()
        
        # 创建文件处理器（带轮转）
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=max_bytes,
            backupCount=backup_count,
            encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(detailed_formatter)
        self.logger.addHandler(file_handler)
        
        # 创建错误日志文件处理器
        error_log_file = os.path.join(log_dir, f"pitchppt_error_{timestamp}.log")
        error_handler = RotatingFileHandler(
            error_log_file,
            maxBytes=max_bytes,
            backupCount=backup_count,
            encoding='utf-8'
        )
        error_handler.setLevel(logging.ERROR)
        error_handler.setFormatter(detailed_formatter)
        self.logger.addHandler(error_handler)
        
        # 创建控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(simple_formatter)
        self.logger.addHandler(console_handler)
        
        # 创建性能日志记录器
        self.performance_logger = PerformanceLogger(self.logger)
        
        self.logger.info("=" * 60)
        self.logger.info("PitchPPT 日志系统初始化完成")
        self.logger.info(f"日志文件: {log_file}")
        self.logger.info(f"错误日志: {error_log_file}")
        self.logger.info("=" * 60)
    
    def get_logger(self) -> logging.Logger:
        """
        获取日志记录器
        
        Returns:
            logging.Logger: 日志记录器实例
        """
        return self.logger
    
    def get_performance_logger(self) -> PerformanceLogger:
        """
        获取性能日志记录器
        
        Returns:
            PerformanceLogger: 性能日志记录器实例
        """
        return self.performance_logger
    
    def set_level(self, level: str):
        """
        设置日志级别
        
        Args:
            level: 日志级别 (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        """
        level_map = {
            'DEBUG': logging.DEBUG,
            'INFO': logging.INFO,
            'WARNING': logging.WARNING,
            'ERROR': logging.ERROR,
            'CRITICAL': logging.CRITICAL
        }
        
        if level.upper() in level_map:
            self.logger.setLevel(level_map[level.upper()])
            self.logger.info(f"日志级别已设置为: {level.upper()}")
        else:
            self.logger.warning(f"无效的日志级别: {level}")
    
    def log_conversion_start(self, input_path: str, output_path: str, mode: str):
        """
        记录转换开始
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            mode: 转换模式
        """
        self.logger.info("=" * 60)
        self.logger.info("开始PPT转换任务")
        self.logger.info(f"输入文件: {input_path}")
        self.logger.info(f"输出文件: {output_path}")
        self.logger.info(f"转换模式: {mode}")
        self.logger.info("=" * 60)
    
    def log_conversion_end(self, success: bool, output_path: str, duration: float = None):
        """
        记录转换结束
        
        Args:
            success: 是否成功
            output_path: 输出文件路径
            duration: 耗时（秒）
        """
        self.logger.info("=" * 60)
        if success:
            self.logger.info("转换成功完成")
            if duration:
                self.logger.info(f"总耗时: {duration:.3f}秒")
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path) / 1024 / 1024
                self.logger.info(f"输出文件大小: {file_size:.2f} MB")
        else:
            self.logger.error("转换失败")
        self.logger.info("=" * 60)
    
    def log_error_with_context(self, error: Exception, context: str = ""):
        """
        记录带上下文的错误
        
        Args:
            error: 异常对象
            context: 上下文信息
        """
        self.logger.error(f"发生错误: {context}")
        self.logger.error(f"错误类型: {type(error).__name__}")
        self.logger.error(f"错误信息: {str(error)}")
        self.logger.exception("详细堆栈信息:")
    
    def cleanup_old_logs(self, days: int = 30):
        """
        清理旧日志文件
        
        Args:
            days: 保留天数
        """
        if not os.path.exists(self.log_dir):
            return
        
        current_time = time.time()
        cutoff_time = current_time - (days * 24 * 60 * 60)
        
        deleted_count = 0
        for filename in os.listdir(self.log_dir):
            file_path = os.path.join(self.log_dir, filename)
            if os.path.isfile(file_path) and filename.endswith('.log'):
                file_time = os.path.getmtime(file_path)
                if file_time < cutoff_time:
                    try:
                        os.remove(file_path)
                        deleted_count += 1
                        self.logger.info(f"已删除旧日志文件: {filename}")
                    except Exception as e:
                        self.logger.warning(f"删除日志文件失败 {filename}: {e}")
        
        if deleted_count > 0:
            self.logger.info(f"共清理了 {deleted_count} 个旧日志文件")