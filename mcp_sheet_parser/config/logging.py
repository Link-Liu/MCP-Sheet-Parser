# logging.py
# 日志相关配置

from dataclasses import dataclass
from typing import Dict, List, Optional
import logging


@dataclass
class LoggingConfig:
    """日志配置"""
    
    # 基本配置
    LOG_LEVEL: str = 'INFO'
    LOG_FORMAT: str = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    DATE_FORMAT: str = '%Y-%m-%d %H:%M:%S'
    
    # 文件日志配置
    ENABLE_FILE_LOGGING: bool = False
    LOG_FILE_PATH: str = 'mcp_sheet_parser.log'
    LOG_FILE_MAX_SIZE: int = 10  # MB
    LOG_FILE_BACKUP_COUNT: int = 5
    
    # 控制台日志配置
    ENABLE_CONSOLE_LOGGING: bool = True
    CONSOLE_LOG_LEVEL: str = None  # 如果为None，使用LOG_LEVEL
    
    # 结构化日志配置
    ENABLE_STRUCTURED_LOGGING: bool = False
    STRUCTURED_FORMAT: str = 'json'  # json, yaml
    
    # 日志过滤配置
    LOG_FILTERS: Dict[str, str] = None
    EXCLUDE_MODULES: List[str] = None
    
    # 性能日志配置
    ENABLE_PERFORMANCE_LOGGING: bool = True
    PERFORMANCE_LOG_THRESHOLD: float = 1.0  # 秒
    
    # 错误日志配置
    ENABLE_ERROR_TRACKING: bool = True
    ERROR_LOG_DETAIL_LEVEL: str = 'full'  # basic, full, minimal
    
    def __post_init__(self):
        """初始化默认值"""
        if self.CONSOLE_LOG_LEVEL is None:
            self.CONSOLE_LOG_LEVEL = self.LOG_LEVEL
        
        if self.LOG_FILTERS is None:
            self.LOG_FILTERS = {
                'urllib3': 'WARNING',  # 减少HTTP库日志
                'requests': 'WARNING',
                'matplotlib': 'WARNING',
                'PIL': 'WARNING'
            }
        
        if self.EXCLUDE_MODULES is None:
            self.EXCLUDE_MODULES = [
                '__pycache__',
                'test_',
                'debug_'
            ]
    
    def get_log_level(self, level_name: str = None) -> int:
        """获取日志级别"""
        level_name = level_name or self.LOG_LEVEL
        level_map = {
            'DEBUG': logging.DEBUG,
            'INFO': logging.INFO,
            'WARNING': logging.WARNING,
            'ERROR': logging.ERROR,
            'CRITICAL': logging.CRITICAL
        }
        return level_map.get(level_name.upper(), logging.INFO)
    
    def get_formatter(self, structured: bool = None) -> logging.Formatter:
        """获取日志格式化器"""
        if structured is None:
            structured = self.ENABLE_STRUCTURED_LOGGING
        
        if structured:
            return self._get_structured_formatter()
        else:
            return logging.Formatter(
                fmt=self.LOG_FORMAT,
                datefmt=self.DATE_FORMAT
            )
    
    def _get_structured_formatter(self) -> logging.Formatter:
        """获取结构化日志格式化器"""
        if self.STRUCTURED_FORMAT == 'json':
            return JsonFormatter()
        else:
            # 简化的结构化格式
            return logging.Formatter(
                fmt='%(asctime)s | %(levelname)s | %(name)s | %(message)s',
                datefmt=self.DATE_FORMAT
            )
    
    def get_console_handler(self) -> logging.StreamHandler:
        """获取控制台处理器"""
        handler = logging.StreamHandler()
        handler.setLevel(self.get_log_level(self.CONSOLE_LOG_LEVEL))
        handler.setFormatter(self.get_formatter())
        return handler
    
    def get_file_handler(self) -> Optional[logging.Handler]:
        """获取文件处理器"""
        if not self.ENABLE_FILE_LOGGING:
            return None
        
        try:
            from logging.handlers import RotatingFileHandler
            handler = RotatingFileHandler(
                filename=self.LOG_FILE_PATH,
                maxBytes=self.LOG_FILE_MAX_SIZE * 1024 * 1024,
                backupCount=self.LOG_FILE_BACKUP_COUNT,
                encoding='utf-8'
            )
            handler.setLevel(self.get_log_level())
            handler.setFormatter(self.get_formatter(structured=True))
            return handler
        except Exception:
            # 如果文件处理器创建失败，返回None
            return None
    
    def configure_logger(self, logger_name: str = None) -> logging.Logger:
        """配置日志器"""
        logger = logging.getLogger(logger_name)
        logger.setLevel(self.get_log_level())
        
        # 清除已有处理器
        logger.handlers.clear()
        
        # 添加控制台处理器
        if self.ENABLE_CONSOLE_LOGGING:
            console_handler = self.get_console_handler()
            logger.addHandler(console_handler)
        
        # 添加文件处理器
        if self.ENABLE_FILE_LOGGING:
            file_handler = self.get_file_handler()
            if file_handler:
                logger.addHandler(file_handler)
        
        # 应用过滤器
        self._apply_filters(logger)
        
        return logger
    
    def _apply_filters(self, logger: logging.Logger):
        """应用日志过滤器"""
        for module_name, level in self.LOG_FILTERS.items():
            module_logger = logging.getLogger(module_name)
            module_logger.setLevel(self.get_log_level(level))
    
    def get_performance_logger_config(self) -> Dict[str, any]:
        """获取性能日志配置"""
        return {
            'enabled': self.ENABLE_PERFORMANCE_LOGGING,
            'threshold': self.PERFORMANCE_LOG_THRESHOLD,
            'level': 'INFO'
        }
    
    def optimize_for_production(self) -> 'LoggingConfig':
        """为生产环境优化配置"""
        optimized = LoggingConfig(
            LOG_LEVEL='WARNING',  # 只记录警告和错误
            LOG_FORMAT=self.LOG_FORMAT,
            DATE_FORMAT=self.DATE_FORMAT,
            ENABLE_FILE_LOGGING=True,  # 启用文件日志
            LOG_FILE_PATH=self.LOG_FILE_PATH,
            LOG_FILE_MAX_SIZE=self.LOG_FILE_MAX_SIZE,
            LOG_FILE_BACKUP_COUNT=self.LOG_FILE_BACKUP_COUNT,
            ENABLE_CONSOLE_LOGGING=False,  # 禁用控制台日志
            CONSOLE_LOG_LEVEL=self.CONSOLE_LOG_LEVEL,
            ENABLE_STRUCTURED_LOGGING=True,  # 启用结构化日志
            STRUCTURED_FORMAT=self.STRUCTURED_FORMAT,
            LOG_FILTERS=self.LOG_FILTERS,
            EXCLUDE_MODULES=self.EXCLUDE_MODULES,
            ENABLE_PERFORMANCE_LOGGING=False,  # 禁用性能日志
            PERFORMANCE_LOG_THRESHOLD=self.PERFORMANCE_LOG_THRESHOLD,
            ENABLE_ERROR_TRACKING=True,
            ERROR_LOG_DETAIL_LEVEL='full'
        )
        
        return optimized
    
    def optimize_for_development(self) -> 'LoggingConfig':
        """为开发环境优化配置"""
        optimized = LoggingConfig(
            LOG_LEVEL='DEBUG',  # 详细日志
            LOG_FORMAT=self.LOG_FORMAT,
            DATE_FORMAT=self.DATE_FORMAT,
            ENABLE_FILE_LOGGING=False,  # 禁用文件日志
            LOG_FILE_PATH=self.LOG_FILE_PATH,
            LOG_FILE_MAX_SIZE=self.LOG_FILE_MAX_SIZE,
            LOG_FILE_BACKUP_COUNT=self.LOG_FILE_BACKUP_COUNT,
            ENABLE_CONSOLE_LOGGING=True,  # 启用控制台日志
            CONSOLE_LOG_LEVEL='DEBUG',
            ENABLE_STRUCTURED_LOGGING=False,  # 禁用结构化日志
            STRUCTURED_FORMAT=self.STRUCTURED_FORMAT,
            LOG_FILTERS={},  # 不过滤任何模块
            EXCLUDE_MODULES=self.EXCLUDE_MODULES,
            ENABLE_PERFORMANCE_LOGGING=True,  # 启用性能日志
            PERFORMANCE_LOG_THRESHOLD=0.1,  # 更低的阈值
            ENABLE_ERROR_TRACKING=True,
            ERROR_LOG_DETAIL_LEVEL='full'
        )
        
        return optimized


class JsonFormatter(logging.Formatter):
    """JSON格式化器"""
    
    def format(self, record: logging.LogRecord) -> str:
        import json
        from datetime import datetime
        
        log_data = {
            'timestamp': datetime.fromtimestamp(record.created).isoformat(),
            'level': record.levelname,
            'logger': record.name,
            'message': record.getMessage(),
            'module': record.module,
            'function': record.funcName,
            'line': record.lineno
        }
        
        # 添加异常信息
        if record.exc_info:
            log_data['exception'] = self.formatException(record.exc_info)
        
        # 添加额外字段
        if hasattr(record, 'extra_data'):
            log_data.update(record.extra_data)
        
        return json.dumps(log_data, ensure_ascii=False) 