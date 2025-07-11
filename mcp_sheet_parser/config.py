# config.py
# 简化的配置模块，专注于HTML转换

from typing import Dict, Set


class Config:
    """MCP-Sheet-Parser 配置类"""
    
    # 支持的文件格式
    SUPPORTED_EXCEL_FORMATS: Set[str] = {
        '.xlsx', '.xls', '.xlt', '.xlsb', '.xlsm', '.xltm', '.xltx'
    }
    
    SUPPORTED_CSV_FORMATS: Set[str] = {
        '.csv'
    }
    
    SUPPORTED_WPS_FORMATS: Set[str] = {
        '.et', '.ett', '.ets'
    }
    
    @classmethod
    def get_all_supported_formats(cls) -> Set[str]:
        """获取所有支持的文件格式"""
        return cls.SUPPORTED_EXCEL_FORMATS | cls.SUPPORTED_CSV_FORMATS | cls.SUPPORTED_WPS_FORMATS
    
    # 性能限制
    MAX_FILE_SIZE_MB: int = 500  # 提升到500MB
    MAX_ROWS: int = 1000000
    MAX_COLS: int = 16384
    
    # 性能优化配置
    ENABLE_PERFORMANCE_MODE: bool = True    # 启用性能优化模式
    CHUNK_SIZE: int = 1000                  # 分块大小
    MAX_MEMORY_MB: int = 2048              # 最大内存使用
    ENABLE_PROGRESS_TRACKING: bool = True   # 启用进度跟踪
    ENABLE_PARALLEL_PROCESSING: bool = True # 启用并行处理
    
    # HTML输出设置
    HTML_DEFAULT_ENCODING: str = 'utf-8'
    HTML_TABLE_BORDER: int = 1
    HTML_CELL_SPACING: int = 0
    HTML_CELL_PADDING: int = 4
    
    # 默认样式
    DEFAULT_FONT_FAMILY: str = 'Arial, sans-serif'
    DEFAULT_FONT_SIZE: int = 12
    DEFAULT_BORDER_COLOR: str = '#000000'
    DEFAULT_BACKGROUND_COLOR: str = '#FFFFFF'
    DEFAULT_TEXT_COLOR: str = '#000000'
    
    # 边框样式映射
    BORDER_STYLE_MAPPING: Dict[str, str] = {
        'thin': '1px solid',
        'medium': '2px solid',
        'thick': '3px solid',
        'dashed': '1px dashed',
        'dotted': '1px dotted',
        'double': '3px double',
        'hair': '1px solid',
        'mediumDashed': '2px dashed',
        'dashDot': '1px dashed',
        'mediumDashDot': '2px dashed',
        'dashDotDot': '1px dashed',
        'mediumDashDotDot': '2px dashed',
        'slantDashDot': '1px dashed'
    }
    
    # 对齐方式映射
    ALIGNMENT_MAPPING: Dict[str, str] = {
        'left': 'left',
        'center': 'center',
        'right': 'right',
        'justify': 'justify',
        'centerContinuous': 'center',
        'distributed': 'justify',
        'fill': 'left',
        'general': 'left'
    }
    
    VERTICAL_ALIGNMENT_MAPPING: Dict[str, str] = {
        'top': 'top',
        'center': 'middle',
        'bottom': 'bottom',
        'justify': 'middle',
        'distributed': 'middle'
    }
    
    # 输出选项
    INCLUDE_EMPTY_CELLS: bool = True
    INCLUDE_COMMENTS: bool = True
    INCLUDE_HYPERLINKS: bool = True
    
    # 公式处理选项
    ENABLE_FORMULA_PROCESSING: bool = True    # 启用公式处理
    SHOW_FORMULA_TEXT: bool = True           # 显示公式文本
    CALCULATE_FORMULAS: bool = True          # 计算公式结果
    SHOW_FORMULA_ERRORS: bool = True         # 显示公式错误
    FORMULA_CACHE_SIZE: int = 1000          # 公式缓存大小
    SUPPORTED_FUNCTIONS_ONLY: bool = False   # 仅处理支持的函数
    
    # 图表转换配置
    ENABLE_CHART_CONVERSION: bool = True        # 启用图表转换功能
    CHART_OUTPUT_FORMAT: str = 'svg'           # 图表输出格式：svg, png, canvas
    CHART_DEFAULT_WIDTH: int = 600             # 默认图表宽度
    CHART_DEFAULT_HEIGHT: int = 400            # 默认图表高度
    CHART_QUALITY: str = 'high'                # 图表质量：low, medium, high
    SUPPORTED_CHART_TYPES: list = [             # 支持的图表类型
        'column', 'bar', 'line', 'pie', 'area', 'scatter'
    ]
    CHART_COLOR_SCHEME: str = 'default'        # 配色方案：default, business, modern
    CHART_EMBED_MODE: str = 'inline'           # 嵌入模式：inline, external
    CHART_RESPONSIVE: bool = True               # 响应式图表
    CHART_ANIMATIONS: bool = False              # 图表动画（暂不支持）
    
    # 日志设置
    LOG_LEVEL: str = 'INFO'


# 主题样式定义
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
    """错误信息常量"""
    
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
