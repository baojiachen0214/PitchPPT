import traceback
from enum import Enum
from typing import Optional, Dict, Any, Callable
from functools import wraps
from src.utils.logger import Logger

class ErrorSeverity(Enum):
    """
    错误严重程度
    """
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"

class ErrorCategory(Enum):
    """
    错误类别
    """
    FILE_IO = "file_io"
    CONVERSION = "conversion"
    VALIDATION = "validation"
    NETWORK = "network"
    SYSTEM = "system"
    USER_INPUT = "user_input"
    UNKNOWN = "unknown"

class PitchPPTError(Exception):
    """
    PitchPPT基础异常类
    """
    
    def __init__(self, message: str, category: ErrorCategory = ErrorCategory.UNKNOWN,
                 severity: ErrorSeverity = ErrorSeverity.ERROR,
                 details: Dict[str, Any] = None,
                 recoverable: bool = True):
        """
        初始化异常
        
        Args:
            message: 错误消息
            category: 错误类别
            severity: 严重程度
            details: 错误详情
            recoverable: 是否可恢复
        """
        super().__init__(message)
        self.message = message
        self.category = category
        self.severity = severity
        self.details = details or {}
        self.recoverable = recoverable
        self.original_exception = None
    
    def to_dict(self) -> Dict[str, Any]:
        """
        转换为字典
        
        Returns:
            Dict[str, Any]: 错误信息字典
        """
        return {
            'message': self.message,
            'category': self.category.value,
            'severity': self.severity.value,
            'details': self.details,
            'recoverable': self.recoverable,
            'type': self.__class__.__name__
        }
    
    def __str__(self) -> str:
        return f"[{self.category.value.upper()}] {self.message}"

class FileIOError(PitchPPTError):
    """
    文件IO错误
    """
    
    def __init__(self, message: str, file_path: str = None, **kwargs):
        details = kwargs.get('details', {})
        if file_path:
            details['file_path'] = file_path
        kwargs['details'] = details
        kwargs['category'] = ErrorCategory.FILE_IO
        super().__init__(message, **kwargs)

class ConversionError(PitchPPTError):
    """
    转换错误
    """
    
    def __init__(self, message: str, input_file: str = None, output_file: str = None, **kwargs):
        details = kwargs.get('details', {})
        if input_file:
            details['input_file'] = input_file
        if output_file:
            details['output_file'] = output_file
        kwargs['details'] = details
        kwargs['category'] = ErrorCategory.CONVERSION
        super().__init__(message, **kwargs)

class ValidationError(PitchPPTError):
    """
    验证错误
    """
    
    def __init__(self, message: str, field: str = None, value: Any = None, **kwargs):
        details = kwargs.get('details', {})
        if field:
            details['field'] = field
        if value is not None:
            details['value'] = str(value)
        kwargs['details'] = details
        kwargs['category'] = ErrorCategory.VALIDATION
        kwargs['recoverable'] = True
        super().__init__(message, **kwargs)

class SystemError(PitchPPTError):
    """
    系统错误
    """
    
    def __init__(self, message: str, **kwargs):
        kwargs['category'] = ErrorCategory.SYSTEM
        kwargs['recoverable'] = False
        super().__init__(message, **kwargs)

class ErrorHandler:
    """
    错误处理器
    提供统一的错误处理和恢复机制
    """
    
    def __init__(self):
        self.logger = Logger().get_logger()
        self._error_handlers = {}
        self._recovery_strategies = {}
    
    def register_handler(self, error_type: type, handler: Callable[[Exception], bool]):
        """
        注册错误处理器
        
        Args:
            error_type: 异常类型
            handler: 处理函数，返回是否成功处理
        """
        self._error_handlers[error_type] = handler
        self.logger.debug(f"注册错误处理器: {error_type.__name__}")
    
    def register_recovery(self, error_type: type, strategy: Callable[[Exception], bool]):
        """
        注册恢复策略
        
        Args:
            error_type: 异常类型
            strategy: 恢复策略函数，返回是否成功恢复
        """
        self._recovery_strategies[error_type] = strategy
        self.logger.debug(f"注册恢复策略: {error_type.__name__}")
    
    def handle_error(self, error: Exception, context: str = "") -> bool:
        """
        处理错误
        
        Args:
            error: 异常对象
            context: 上下文信息
            
        Returns:
            bool: 是否成功处理
        """
        error_type = type(error)
        
        # 记录错误
        self._log_error(error, context)
        
        # 尝试使用注册的处理器
        if error_type in self._error_handlers:
            try:
                if self._error_handlers[error_type](error):
                    self.logger.info(f"错误处理器成功处理: {error_type.__name__}")
                    return True
            except Exception as e:
                self.logger.error(f"错误处理器执行失败: {e}")
        
        # 尝试恢复
        if error_type in self._recovery_strategies:
            try:
                if self._recovery_strategies[error_type](error):
                    self.logger.info(f"恢复策略成功执行: {error_type.__name__}")
                    return True
            except Exception as e:
                self.logger.error(f"恢复策略执行失败: {e}")
        
        return False
    
    def _log_error(self, error: Exception, context: str):
        """
        记录错误
        
        Args:
            error: 异常对象
            context: 上下文信息
        """
        if isinstance(error, PitchPPTError):
            if error.severity == ErrorSeverity.CRITICAL:
                self.logger.critical(f"{context}: {error}")
            elif error.severity == ErrorSeverity.ERROR:
                self.logger.error(f"{context}: {error}")
            elif error.severity == ErrorSeverity.WARNING:
                self.logger.warning(f"{context}: {error}")
            else:
                self.logger.info(f"{context}: {error}")
            
            if error.details:
                self.logger.debug(f"错误详情: {error.details}")
        else:
            self.logger.error(f"{context}: {type(error).__name__}: {str(error)}")
        
        # 记录堆栈跟踪
        if not isinstance(error, PitchPPTError) or error.severity in [ErrorSeverity.ERROR, ErrorSeverity.CRITICAL]:
            self.logger.debug(f"堆栈跟踪:\n{traceback.format_exc()}")
    
    def get_user_friendly_message(self, error: Exception) -> str:
        """
        获取用户友好的错误消息
        
        Args:
            error: 异常对象
            
        Returns:
            str: 用户友好的错误消息
        """
        if isinstance(error, PitchPPTError):
            return error.message
        
        error_messages = {
            FileNotFoundError: "找不到指定的文件，请检查文件路径是否正确。",
            PermissionError: "没有权限访问该文件，请检查文件权限。",
            ValueError: "输入的值无效，请检查您的输入。",
            TypeError: "数据类型错误，请检查您的输入。",
            MemoryError: "内存不足，请关闭其他程序后重试。",
            TimeoutError: "操作超时，请稍后重试。",
            ConnectionError: "网络连接失败，请检查网络设置。",
        }
        
        error_type = type(error)
        return error_messages.get(error_type, f"发生未知错误: {str(error)}")
    
    def wrap_exception(self, original_error: Exception, new_message: str,
                      new_category: ErrorCategory = None,
                      new_severity: ErrorSeverity = None) -> PitchPPTError:
        """
        包装原始异常
        
        Args:
            original_error: 原始异常
            new_message: 新的错误消息
            new_category: 新的错误类别
            new_severity: 新的严重程度
            
        Returns:
            PitchPPTError: 包装后的异常
        """
        if isinstance(original_error, PitchPPTError):
            pitch_error = original_error
        else:
            pitch_error = PitchPPTError(str(original_error))
        
        if new_message:
            pitch_error.message = new_message
        if new_category:
            pitch_error.category = new_category
        if new_severity:
            pitch_error.severity = new_severity
        
        pitch_error.original_exception = original_error
        return pitch_error

def handle_errors(error_handler: ErrorHandler = None, default_message: str = None,
                reraise: bool = False):
    """
    错误处理装饰器
    
    Args:
        error_handler: 错误处理器实例
        default_message: 默认错误消息
        reraise: 是否重新抛出异常
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            handler = error_handler or ErrorHandler()
            logger = Logger().get_logger()
            
            try:
                return func(*args, **kwargs)
            except PitchPPTError as e:
                if not handler.handle_error(e, f"{func.__name__}"):
                    logger.error(f"无法处理的PitchPPT错误: {e}")
                    if reraise:
                        raise
                    return None
            except Exception as e:
                wrapped_error = handler.wrap_exception(e, default_message or f"{func.__name__} 执行失败")
                if not handler.handle_error(wrapped_error, f"{func.__name__}"):
                    logger.error(f"无法处理的异常: {e}")
                    if reraise:
                        raise
                    return None
        
        return wrapper
    return decorator

def retry_on_failure(max_retries: int = 3, delay: float = 1.0,
                    backoff_factor: float = 2.0, allowed_exceptions: tuple = (Exception,)):
    """
    失败重试装饰器
    
    Args:
        max_retries: 最大重试次数
        delay: 初始延迟（秒）
        backoff_factor: 退避因子
        allowed_exceptions: 允许重试的异常类型
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            logger = Logger().get_logger()
            current_delay = delay
            
            for attempt in range(max_retries + 1):
                try:
                    return func(*args, **kwargs)
                except allowed_exceptions as e:
                    if attempt == max_retries:
                        logger.error(f"{func.__name__} 重试 {max_retries} 次后仍然失败")
                        raise
                    
                    logger.warning(f"{func.__name__} 第 {attempt + 1} 次尝试失败: {e}, {current_delay}秒后重试...")
                    import time
                    time.sleep(current_delay)
                    current_delay *= backoff_factor
        
        return wrapper
    return decorator

class ErrorContext:
    """
    错误上下文管理器
    用于捕获和处理特定代码块中的错误
    """
    
    def __init__(self, error_handler: ErrorHandler = None,
                 context_name: str = "", reraise: bool = False):
        """
        初始化错误上下文
        
        Args:
            error_handler: 错误处理器
            context_name: 上下文名称
            reraise: 是否重新抛出异常
        """
        self.error_handler = error_handler or ErrorHandler()
        self.context_name = context_name
        self.reraise = reraise
        self.error = None
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None:
            self.error = exc_val
            if not self.error_handler.handle_error(exc_val, self.context_name):
                if self.reraise:
                    return False
            return True
        return False
    
    def get_error(self) -> Optional[Exception]:
        """
        获取捕获的异常
        
        Returns:
            Optional[Exception]: 异常对象，如果没有异常则返回None
        """
        return self.error
    
    def has_error(self) -> bool:
        """
        是否有错误发生
        
        Returns:
            bool: 是否有错误
        """
        return self.error is not None