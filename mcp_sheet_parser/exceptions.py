# exceptions.py
# 统一异常处理模块 - 自定义异常类、错误处理装饰器和恢复机制

import logging
import traceback
import functools
from typing import Dict, Any, Optional, Union, Callable, Type
from dataclasses import dataclass
from enum import Enum


class ErrorSeverity(Enum):
    """错误严重程度"""
    LOW = "low"           # 轻微错误，可以继续
    MEDIUM = "medium"     # 中等错误，需要注意
    HIGH = "high"         # 严重错误，影响功能
    CRITICAL = "critical" # 致命错误，必须停止


@dataclass
class ErrorContext:
    """错误上下文信息"""
    operation: str                    # 当前操作
    file_path: Optional[str] = None   # 文件路径
    sheet_name: Optional[str] = None  # 工作表名
    cell_position: Optional[str] = None  # 单元格位置
    additional_info: Optional[Dict[str, Any]] = None  # 额外信息
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'operation': self.operation,
            'file_path': self.file_path,
            'sheet_name': self.sheet_name, 
            'cell_position': self.cell_position,
            'additional_info': self.additional_info or {}
        }


class MCPSheetParserError(Exception):
    """MCP-Sheet-Parser 基础异常类"""
    
    def __init__(self, 
                 message: str, 
                 severity: ErrorSeverity = ErrorSeverity.MEDIUM,
                 context: Optional[ErrorContext] = None,
                 original_error: Optional[Exception] = None):
        super().__init__(message)
        self.message = message
        self.severity = severity
        self.context = context
        self.original_error = original_error
        self.timestamp = None
        
    def get_detailed_message(self) -> str:
        """获取详细的错误消息"""
        parts = [f"[{self.severity.value.upper()}] {self.message}"]
        
        if self.context:
            if self.context.file_path:
                parts.append(f"文件: {self.context.file_path}")
            if self.context.sheet_name:
                parts.append(f"工作表: {self.context.sheet_name}")
            if self.context.cell_position:
                parts.append(f"位置: {self.context.cell_position}")
            if self.context.operation:
                parts.append(f"操作: {self.context.operation}")
        
        if self.original_error:
            parts.append(f"原始错误: {type(self.original_error).__name__}: {self.original_error}")
        
        return " | ".join(parts)
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式"""
        original_error_str = None
        if self.original_error:
            original_error_str = f"{type(self.original_error).__name__}: {self.original_error}"
        
        return {
            'error_type': self.__class__.__name__,
            'message': self.message,
            'severity': self.severity.value,
            'context': self.context.to_dict() if self.context else None,
            'original_error': original_error_str,
            'timestamp': self.timestamp
        }


class FileProcessingError(MCPSheetParserError):
    """文件处理错误"""
    pass


class FileNotFoundError(FileProcessingError):
    """文件未找到错误"""
    
    def __init__(self, file_path: str, **kwargs):
        context = kwargs.get('context') or ErrorContext(operation="文件查找")
        context.file_path = file_path
        super().__init__(
            f"文件不存在: {file_path}",
            severity=ErrorSeverity.HIGH,
            context=context,
            **{k: v for k, v in kwargs.items() if k != 'context'}
        )


class UnsupportedFormatError(FileProcessingError):
    """不支持的文件格式错误"""
    
    def __init__(self, file_path: str, extension: str, **kwargs):
        context = kwargs.get('context') or ErrorContext(operation="格式检查")
        context.file_path = file_path
        context.additional_info = {'extension': extension}
        super().__init__(
            f"不支持的文件格式: {extension} (文件: {file_path})",
            severity=ErrorSeverity.HIGH,
            context=context,
            **{k: v for k, v in kwargs.items() if k != 'context'}
        )


class FileSizeExceededError(FileProcessingError):
    """文件大小超限错误"""
    
    def __init__(self, file_path: str, size_mb: float, max_size_mb: int, **kwargs):
        context = kwargs.get('context') or ErrorContext(operation="文件大小检查")
        context.file_path = file_path
        context.additional_info = {'size_mb': size_mb, 'max_size_mb': max_size_mb}
        super().__init__(
            f"文件过大: {size_mb:.2f}MB > {max_size_mb}MB (文件: {file_path})",
            severity=ErrorSeverity.HIGH,
            context=context,
            **{k: v for k, v in kwargs.items() if k != 'context'}
        )


class ParsingError(MCPSheetParserError):
    """解析错误"""
    pass


class WorksheetParsingError(ParsingError):
    """工作表解析错误"""
    
    def __init__(self, sheet_name: str, message: str, **kwargs):
        context = kwargs.get('context') or ErrorContext(operation="工作表解析")
        context.sheet_name = sheet_name
        super().__init__(
            f"工作表解析失败 '{sheet_name}': {message}",
            severity=ErrorSeverity.MEDIUM,
            context=context,
            **{k: v for k, v in kwargs.items() if k != 'context'}
        )


class CellProcessingError(ParsingError):
    """单元格处理错误"""
    
    def __init__(self, position: str, message: str, **kwargs):
        context = kwargs.get('context') or ErrorContext(operation="单元格处理")
        context.cell_position = position
        super().__init__(
            f"单元格处理失败 {position}: {message}",
            severity=ErrorSeverity.LOW,
            context=context,
            **{k: v for k, v in kwargs.items() if k != 'context'}
        )


class ConversionError(MCPSheetParserError):
    """转换错误"""
    pass


class HTMLConversionError(ConversionError):
    """HTML转换错误"""
    
    def __init__(self, message: str, **kwargs):
        context = kwargs.get('context') or ErrorContext(operation="HTML转换")
        super().__init__(
            f"HTML转换失败: {message}",
            severity=ErrorSeverity.MEDIUM,
            context=context,
            **{k: v for k, v in kwargs.items() if k != 'context'}
        )


class FormulaProcessingError(MCPSheetParserError):
    """公式处理错误"""
    
    def __init__(self, formula: str, message: str, **kwargs):
        context = kwargs.get('context') or ErrorContext(operation="公式处理")
        context.additional_info = {'formula': formula}
        super().__init__(
            f"公式处理失败 '{formula}': {message}",
            severity=ErrorSeverity.LOW,
            context=context,
            **{k: v for k, v in kwargs.items() if k != 'context'}
        )


class ChartConversionError(ConversionError):
    """图表转换错误"""
    
    def __init__(self, chart_type: str, message: str, **kwargs):
        context = kwargs.get('context') or ErrorContext(operation="图表转换")
        context.additional_info = {'chart_type': chart_type}
        super().__init__(
            f"图表转换失败 [{chart_type}]: {message}",
            severity=ErrorSeverity.LOW,
            context=context,
            **{k: v for k, v in kwargs.items() if k != 'context'}
        )


class SecurityError(MCPSheetParserError):
    """安全错误"""
    
    def __init__(self, message: str, **kwargs):
        context = kwargs.get('context') or ErrorContext(operation="安全检查")
        super().__init__(
            f"安全检查失败: {message}",
            severity=ErrorSeverity.HIGH,
            context=context,
            **{k: v for k, v in kwargs.items() if k != 'context'}
        )


class PerformanceError(MCPSheetParserError):
    """性能相关错误"""
    
    def __init__(self, message: str, **kwargs):
        context = kwargs.get('context') or ErrorContext(operation="性能优化")
        super().__init__(
            f"性能错误: {message}",
            severity=ErrorSeverity.MEDIUM,
            context=context,
            **{k: v for k, v in kwargs.items() if k != 'context'}
        )


class ConfigurationError(MCPSheetParserError):
    """配置错误"""
    
    def __init__(self, parameter: str, message: str, **kwargs):
        context = kwargs.get('context') or ErrorContext(operation="配置验证")
        context.additional_info = {'parameter': parameter}
        super().__init__(
            f"配置错误 [{parameter}]: {message}",
            severity=ErrorSeverity.MEDIUM,
            context=context,
            **{k: v for k, v in kwargs.items() if k != 'context'}
        )


class ErrorHandler:
    """统一错误处理器"""
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        self.logger = logger or logging.getLogger(__name__)
        self.error_stats = {
            'total_errors': 0,
            'by_type': {},
            'by_severity': {severity.value: 0 for severity in ErrorSeverity}
        }
    
    def handle_error(self, error: Exception, context: Optional[ErrorContext] = None) -> MCPSheetParserError:
        """处理异常并转换为统一的错误格式"""
        
        # 如果已经是我们的自定义异常，直接返回
        if isinstance(error, MCPSheetParserError):
            self._log_error(error)
            self._update_stats(error)
            return error
        
        # 根据原始异常类型创建相应的自定义异常
        if isinstance(error, FileNotFoundError):
            custom_error = FileNotFoundError(
                str(error), context=context, original_error=error
            )
        elif isinstance(error, ValueError):
            custom_error = ParsingError(
                str(error), 
                severity=ErrorSeverity.MEDIUM,
                context=context, 
                original_error=error
            )
        elif isinstance(error, PermissionError):
            custom_error = SecurityError(
                f"权限错误: {error}",
                context=context,
                original_error=error
            )
        elif isinstance(error, MemoryError):
            custom_error = PerformanceError(
                f"内存不足: {error}",
                context=context,
                original_error=error
            )
        else:
            # 通用异常处理
            custom_error = MCPSheetParserError(
                f"未预期的错误: {error}",
                severity=ErrorSeverity.MEDIUM,
                context=context,
                original_error=error
            )
        
        self._log_error(custom_error)
        self._update_stats(custom_error)
        return custom_error
    
    def _log_error(self, error: MCPSheetParserError):
        """记录错误日志"""
        detailed_message = error.get_detailed_message()
        
        if error.severity == ErrorSeverity.CRITICAL:
            self.logger.critical(detailed_message)
        elif error.severity == ErrorSeverity.HIGH:
            self.logger.error(detailed_message)
        elif error.severity == ErrorSeverity.MEDIUM:
            self.logger.warning(detailed_message)
        else:
            self.logger.info(detailed_message)
    
    def _update_stats(self, error: MCPSheetParserError):
        """更新错误统计"""
        self.error_stats['total_errors'] += 1
        
        error_type = type(error).__name__
        if error_type not in self.error_stats['by_type']:
            self.error_stats['by_type'][error_type] = 0
        self.error_stats['by_type'][error_type] += 1
        
        self.error_stats['by_severity'][error.severity.value] += 1
    
    def get_error_summary(self) -> Dict[str, Any]:
        """获取错误统计摘要"""
        return self.error_stats.copy()
    
    def reset_stats(self):
        """重置错误统计"""
        self.error_stats = {
            'total_errors': 0,
            'by_type': {},
            'by_severity': {severity.value: 0 for severity in ErrorSeverity}
        }


def error_handler(
    error_types: Union[Type[Exception], tuple] = Exception,
    severity: ErrorSeverity = ErrorSeverity.MEDIUM,
    operation: str = "未知操作",
    reraise: bool = True
):
    """错误处理装饰器
    
    Args:
        error_types: 要捕获的异常类型
        severity: 错误严重程度
        operation: 操作描述
        reraise: 是否重新抛出异常
    """
    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            handler = ErrorHandler()
            context = ErrorContext(operation=operation)
            
            try:
                return func(*args, **kwargs)
            except error_types as e:
                custom_error = handler.handle_error(e, context)
                if reraise:
                    raise custom_error
                return None
        return wrapper
    return decorator


def safe_execute(
    func: Callable,
    *args,
    error_types: Union[Type[Exception], tuple] = Exception,
    default_value: Any = None,
    operation: str = "安全执行",
    **kwargs
) -> Any:
    """安全执行函数，捕获异常并返回默认值
    
    Args:
        func: 要执行的函数
        error_types: 要捕获的异常类型
        default_value: 异常时返回的默认值
        operation: 操作描述
        
    Returns:
        函数执行结果或默认值
    """
    handler = ErrorHandler()
    context = ErrorContext(operation=operation)
    
    try:
        return func(*args, **kwargs)
    except error_types as e:
        handler.handle_error(e, context)
        return default_value


# 创建全局错误处理器实例
global_error_handler = ErrorHandler()


def create_error_context(operation: str, **kwargs) -> ErrorContext:
    """创建错误上下文的便利函数"""
    return ErrorContext(
        operation=operation,
        file_path=kwargs.get('file_path'),
        sheet_name=kwargs.get('sheet_name'),
        cell_position=kwargs.get('cell_position'),
        additional_info=kwargs.get('additional_info')
    ) 