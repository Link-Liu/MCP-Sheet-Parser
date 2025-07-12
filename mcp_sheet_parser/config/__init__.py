# __init__.py
# 配置模块初始化文件

from .file_format import FileFormatConfig
from .performance import PerformanceConfig
from .html import HTMLConfig
from .style import StyleConfig
from .formula import FormulaConfig
from .chart import ChartConfig
from .logging import LoggingConfig, JsonFormatter

from dataclasses import dataclass
from typing import Dict, Set


@dataclass
class UnifiedConfig:
    """统一配置类 - 聚合所有子配置模块"""
    
    # 各子模块配置
    file_format: FileFormatConfig = None
    performance: PerformanceConfig = None
    html: HTMLConfig = None
    style: StyleConfig = None
    formula: FormulaConfig = None
    chart: ChartConfig = None
    logging: LoggingConfig = None
    
    def __post_init__(self):
        """初始化所有子配置"""
        if self.file_format is None:
            self.file_format = FileFormatConfig()
        
        if self.performance is None:
            self.performance = PerformanceConfig()
        
        if self.html is None:
            self.html = HTMLConfig()
        
        if self.style is None:
            self.style = StyleConfig()
        
        if self.formula is None:
            self.formula = FormulaConfig()
        
        if self.chart is None:
            self.chart = ChartConfig()
        
        if self.logging is None:
            self.logging = LoggingConfig()
    
    # 向后兼容的属性映射
    @property
    def SUPPORTED_EXCEL_FORMATS(self) -> Set[str]:
        return self.file_format.SUPPORTED_EXCEL_FORMATS
    
    @property
    def SUPPORTED_CSV_FORMATS(self) -> Set[str]:
        return self.file_format.SUPPORTED_CSV_FORMATS
    
    @property
    def SUPPORTED_WPS_FORMATS(self) -> Set[str]:
        return self.file_format.SUPPORTED_WPS_FORMATS
    
    def get_all_supported_formats(self) -> Set[str]:
        return self.file_format.get_all_supported_formats()
    
    @property
    def MAX_FILE_SIZE_MB(self) -> int:
        return self.performance.MAX_FILE_SIZE_MB
    
    @MAX_FILE_SIZE_MB.setter
    def MAX_FILE_SIZE_MB(self, value: int):
        self.performance.MAX_FILE_SIZE_MB = value
    
    @property
    def MAX_ROWS(self) -> int:
        return self.performance.MAX_ROWS
    
    @property
    def MAX_COLS(self) -> int:
        return self.performance.MAX_COLS
    
    @property
    def CHUNK_SIZE(self) -> int:
        return self.performance.CHUNK_SIZE
    
    @CHUNK_SIZE.setter
    def CHUNK_SIZE(self, value: int):
        self.performance.CHUNK_SIZE = value
    
    @property
    def MAX_MEMORY_MB(self) -> int:
        return self.performance.MAX_MEMORY_MB
    
    @MAX_MEMORY_MB.setter
    def MAX_MEMORY_MB(self, value: int):
        self.performance.MAX_MEMORY_MB = value
    
    @property
    def ENABLE_PERFORMANCE_MODE(self) -> bool:
        return self.performance.ENABLE_PERFORMANCE_MODE
    
    @ENABLE_PERFORMANCE_MODE.setter
    def ENABLE_PERFORMANCE_MODE(self, value: bool):
        self.performance.ENABLE_PERFORMANCE_MODE = value
    
    @property
    def ENABLE_PROGRESS_TRACKING(self) -> bool:
        return self.performance.ENABLE_PROGRESS_TRACKING
    
    @ENABLE_PROGRESS_TRACKING.setter
    def ENABLE_PROGRESS_TRACKING(self, value: bool):
        self.performance.ENABLE_PROGRESS_TRACKING = value
    
    @property
    def ENABLE_PARALLEL_PROCESSING(self) -> bool:
        return self.performance.ENABLE_PARALLEL_PROCESSING
    
    @ENABLE_PARALLEL_PROCESSING.setter
    def ENABLE_PARALLEL_PROCESSING(self, value: bool):
        self.performance.ENABLE_PARALLEL_PROCESSING = value
    
    @property
    def HTML_DEFAULT_ENCODING(self) -> str:
        return self.html.HTML_DEFAULT_ENCODING
    
    @property
    def HTML_TABLE_BORDER(self) -> int:
        return self.html.HTML_TABLE_BORDER
    
    @property
    def HTML_CELL_SPACING(self) -> int:
        return self.html.HTML_CELL_SPACING
    
    @property
    def HTML_CELL_PADDING(self) -> int:
        return self.html.HTML_CELL_PADDING
    
    @property
    def INCLUDE_EMPTY_CELLS(self) -> bool:
        return self.html.INCLUDE_EMPTY_CELLS
    
    @INCLUDE_EMPTY_CELLS.setter
    def INCLUDE_EMPTY_CELLS(self, value: bool):
        self.html.INCLUDE_EMPTY_CELLS = value
    
    @property
    def INCLUDE_COMMENTS(self) -> bool:
        return self.html.INCLUDE_COMMENTS
    
    @INCLUDE_COMMENTS.setter
    def INCLUDE_COMMENTS(self, value: bool):
        self.html.INCLUDE_COMMENTS = value
    
    @property
    def INCLUDE_HYPERLINKS(self) -> bool:
        return self.html.INCLUDE_HYPERLINKS
    
    @INCLUDE_HYPERLINKS.setter
    def INCLUDE_HYPERLINKS(self, value: bool):
        self.html.INCLUDE_HYPERLINKS = value
    
    @property
    def DEFAULT_FONT_FAMILY(self) -> str:
        return self.style.DEFAULT_FONT_FAMILY
    
    @property
    def DEFAULT_FONT_SIZE(self) -> int:
        return self.style.DEFAULT_FONT_SIZE
    
    @property
    def DEFAULT_BORDER_COLOR(self) -> str:
        return self.style.DEFAULT_BORDER_COLOR
    
    @property
    def DEFAULT_BACKGROUND_COLOR(self) -> str:
        return self.style.DEFAULT_BACKGROUND_COLOR
    
    @property
    def DEFAULT_TEXT_COLOR(self) -> str:
        return self.style.DEFAULT_TEXT_COLOR
    
    @property
    def BORDER_STYLE_MAPPING(self) -> Dict[str, str]:
        return self.style.BORDER_STYLE_MAPPING
    
    @property
    def ALIGNMENT_MAPPING(self) -> Dict[str, str]:
        return self.style.ALIGNMENT_MAPPING
    
    @property
    def VERTICAL_ALIGNMENT_MAPPING(self) -> Dict[str, str]:
        return self.style.VERTICAL_ALIGNMENT_MAPPING
    
    @property
    def ENABLE_FORMULA_PROCESSING(self) -> bool:
        return self.formula.ENABLE_FORMULA_PROCESSING
    
    @ENABLE_FORMULA_PROCESSING.setter
    def ENABLE_FORMULA_PROCESSING(self, value: bool):
        self.formula.ENABLE_FORMULA_PROCESSING = value
    
    @property
    def SHOW_FORMULA_TEXT(self) -> bool:
        return self.formula.SHOW_FORMULA_TEXT
    
    @SHOW_FORMULA_TEXT.setter
    def SHOW_FORMULA_TEXT(self, value: bool):
        self.formula.SHOW_FORMULA_TEXT = value
    
    @property
    def CALCULATE_FORMULAS(self) -> bool:
        return self.formula.CALCULATE_FORMULAS
    
    @CALCULATE_FORMULAS.setter
    def CALCULATE_FORMULAS(self, value: bool):
        self.formula.CALCULATE_FORMULAS = value
    
    @property
    def SHOW_FORMULA_ERRORS(self) -> bool:
        return self.formula.SHOW_FORMULA_ERRORS
    
    @SHOW_FORMULA_ERRORS.setter
    def SHOW_FORMULA_ERRORS(self, value: bool):
        self.formula.SHOW_FORMULA_ERRORS = value
    
    @property
    def FORMULA_CACHE_SIZE(self) -> int:
        return self.formula.FORMULA_CACHE_SIZE
    
    @property
    def SUPPORTED_FUNCTIONS_ONLY(self) -> bool:
        return self.formula.SUPPORTED_FUNCTIONS_ONLY
    
    @SUPPORTED_FUNCTIONS_ONLY.setter
    def SUPPORTED_FUNCTIONS_ONLY(self, value: bool):
        self.formula.SUPPORTED_FUNCTIONS_ONLY = value
    
    @property
    def ENABLE_CHART_CONVERSION(self) -> bool:
        return self.chart.ENABLE_CHART_CONVERSION
    
    @ENABLE_CHART_CONVERSION.setter
    def ENABLE_CHART_CONVERSION(self, value: bool):
        self.chart.ENABLE_CHART_CONVERSION = value
    
    @property
    def CHART_OUTPUT_FORMAT(self) -> str:
        return self.chart.CHART_OUTPUT_FORMAT
    
    @CHART_OUTPUT_FORMAT.setter
    def CHART_OUTPUT_FORMAT(self, value: str):
        self.chart.CHART_OUTPUT_FORMAT = value
    
    @property
    def CHART_DEFAULT_WIDTH(self) -> int:
        return self.chart.CHART_DEFAULT_WIDTH
    
    @CHART_DEFAULT_WIDTH.setter
    def CHART_DEFAULT_WIDTH(self, value: int):
        self.chart.CHART_DEFAULT_WIDTH = value
    
    @property
    def CHART_DEFAULT_HEIGHT(self) -> int:
        return self.chart.CHART_DEFAULT_HEIGHT
    
    @CHART_DEFAULT_HEIGHT.setter
    def CHART_DEFAULT_HEIGHT(self, value: int):
        self.chart.CHART_DEFAULT_HEIGHT = value
    
    @property
    def CHART_QUALITY(self) -> str:
        return self.chart.CHART_QUALITY
    
    @CHART_QUALITY.setter
    def CHART_QUALITY(self, value: str):
        self.chart.CHART_QUALITY = value
    
    @property
    def SUPPORTED_CHART_TYPES(self) -> list:
        return self.chart.SUPPORTED_CHART_TYPES
    
    @property
    def CHART_COLOR_SCHEME(self) -> str:
        return self.chart.CHART_COLOR_SCHEME
    
    @property
    def CHART_EMBED_MODE(self) -> str:
        return self.chart.CHART_EMBED_MODE
    
    @property
    def CHART_RESPONSIVE(self) -> bool:
        return self.chart.CHART_RESPONSIVE
    
    @CHART_RESPONSIVE.setter
    def CHART_RESPONSIVE(self, value: bool):
        self.chart.CHART_RESPONSIVE = value
    
    @property
    def CHART_ANIMATIONS(self) -> bool:
        return self.chart.CHART_ANIMATIONS
    
    @property
    def LOG_LEVEL(self) -> str:
        return self.logging.LOG_LEVEL
    
    def optimize_for_performance(self) -> 'UnifiedConfig':
        """为性能优化整体配置"""
        return UnifiedConfig(
            file_format=self.file_format,
            performance=self.performance.optimize_for_file_size(100),  # 假设100MB文件
            html=self.html.optimize_for_size(),
            style=self.style,
            formula=self.formula.optimize_for_performance(),
            chart=self.chart.optimize_for_web(),
            logging=self.logging.optimize_for_production()
        )
    
    def optimize_for_quality(self) -> 'UnifiedConfig':
        """为质量优化整体配置"""
        return UnifiedConfig(
            file_format=self.file_format,
            performance=self.performance,
            html=self.html,  # 保留所有HTML功能
            style=self.style,
            formula=self.formula.optimize_for_accuracy(),
            chart=self.chart.optimize_for_print(),
            logging=self.logging.optimize_for_development()
        )
    
    def get_module_config(self, module_name: str):
        """获取特定模块的配置"""
        module_map = {
            'file_format': self.file_format,
            'performance': self.performance,
            'html': self.html,
            'style': self.style,
            'formula': self.formula,
            'chart': self.chart,
            'logging': self.logging
        }
        return module_map.get(module_name)
    
    def update_module_config(self, module_name: str, config):
        """更新特定模块的配置"""
        if module_name == 'file_format':
            self.file_format = config
        elif module_name == 'performance':
            self.performance = config
        elif module_name == 'html':
            self.html = config
        elif module_name == 'style':
            self.style = config
        elif module_name == 'formula':
            self.formula = config
        elif module_name == 'chart':
            self.chart = config
        elif module_name == 'logging':
            self.logging = config


# 为了向后兼容，提供原Config类的别名
Config = UnifiedConfig

# 导出所有配置类
__all__ = [
    'FileFormatConfig',
    'PerformanceConfig', 
    'HTMLConfig',
    'StyleConfig',
    'FormulaConfig',
    'ChartConfig',
    'LoggingConfig',
    'JsonFormatter',
    'UnifiedConfig',
    'Config',
    'THEMES',
    'ErrorMessages'
]

# 导入主题和错误信息以保持向后兼容
THEMES = {
    'default': {
        'name': '默认主题',
        'description': '标准样式，平衡美观与实用',
        'body_style': 'font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5;',
        'table_style': 'border-collapse: collapse; width: 100%; background-color: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1);',
        'cell_style': 'border: 1px solid #ddd; padding: 8px; text-align: left;',
        'header_style': 'background-color: #f8f9fa; font-weight: bold; color: #333;'
    },
    'minimal': {
        'name': '极简主题',
        'description': '极简设计，清爽干净',
        'body_style': 'font-family: "Helvetica Neue", Arial, sans-serif; margin: 20px; background-color: white;',
        'table_style': 'border-collapse: collapse; width: 100%;',
        'cell_style': 'border: 1px solid #e0e0e0; padding: 12px; text-align: left;',
        'header_style': 'border-bottom: 2px solid #333; font-weight: 500; color: #333;'
    },
    'dark': {
        'name': '暗色主题',
        'description': '暗色主题，护眼舒适',
        'body_style': 'font-family: Arial, sans-serif; margin: 20px; background-color: #1a1a1a; color: #e0e0e0;',
        'table_style': 'border-collapse: collapse; width: 100%; background-color: #2d2d2d;',
        'cell_style': 'border: 1px solid #444; padding: 8px; text-align: left; color: #e0e0e0;',
        'header_style': 'background-color: #404040; font-weight: bold; color: #fff;'
    },
    'print': {
        'name': '打印主题',
        'description': '打印优化，黑白适配',
        'body_style': 'font-family: "Times New Roman", serif; margin: 0; background-color: white;',
        'table_style': 'border-collapse: collapse; width: 100%;',
        'cell_style': 'border: 1px solid black; padding: 4px; text-align: left;',
        'header_style': 'background-color: #f0f0f0; font-weight: bold; border: 2px solid black;'
    }
}


class ErrorMessages:
    """错误信息常量 - 保留向后兼容性"""
    
    # 文件相关错误
    FILE_NOT_FOUND: str = "文件不存在: {file_path}"
    FILE_NOT_READABLE: str = "文件不可读: {file_path}"
    FILE_TOO_LARGE: str = "文件过大 ({size}MB > {max_size}MB): {file_path}"
    UNSUPPORTED_FORMAT: str = "不支持的文件格式: {extension}"
    
    # 解析相关错误
    PARSE_ERROR: str = "解析文件时出错: {error}"
    EMPTY_FILE: str = "文件为空: {file_path}"
    CORRUPTED_FILE: str = "文件损坏或格式错误: {file_path}"
    TOO_MANY_ROWS: str = "行数超过限制 ({rows} > {max_rows})"
    TOO_MANY_COLS: str = "列数超过限制 ({cols} > {max_cols})"
    
    # 输出相关错误
    OUTPUT_DIR_ERROR: str = "无法创建输出目录: {dir_path}"
    WRITE_ERROR: str = "写入文件失败: {file_path}"
    PERMISSION_ERROR: str = "没有权限写入文件: {file_path}"
    
    # 一般错误
    UNKNOWN_ERROR: str = "未知错误: {error}"
    INVALID_PARAMETER: str = "无效参数: {parameter}" 

# ====== 向后兼容的配置工厂函数 ======
def create_config_for_performance():
    """性能优化配置工厂"""
    return UnifiedConfig().optimize_for_performance()

def create_config_for_quality():
    """质量优化配置工厂"""
    return UnifiedConfig().optimize_for_quality()

def create_config_for_web():
    """Web优化配置工厂"""
    config = UnifiedConfig()
    config.chart = config.chart.optimize_for_web()
    config.html = config.html.optimize_for_size()
    return config

def create_config_for_print():
    """打印优化配置工厂"""
    config = UnifiedConfig()
    config.chart = config.chart.optimize_for_print()
    return config

def create_minimal_config():
    """极简配置工厂（关闭公式和图表）"""
    config = UnifiedConfig()
    config.formula.ENABLE_FORMULA_PROCESSING = False
    config.chart.ENABLE_CHART_CONVERSION = False
    return config

def create_development_config():
    """开发环境配置工厂"""
    config = UnifiedConfig()
    config.logging = config.logging.optimize_for_development()
    return config

def create_production_config():
    """生产环境配置工厂"""
    config = UnifiedConfig()
    config.logging = config.logging.optimize_for_production()
    return config 