# cli.py
# 命令行接口模块 - 参数解析、主题管理、信息显示

import argparse
import os
from pathlib import Path
from typing import Dict, List, Any

from .config import Config, THEMES
from .utils import get_file_info, setup_logger
from .style_manager import ConditionalRule, ConditionalType, ComparisonOperator


class CLIManager:
    """命令行接口管理器"""
    
    def __init__(self):
        self.logger = setup_logger(__name__)
    
    def create_argument_parser(self) -> argparse.ArgumentParser:
        """创建命令行参数解析器"""
        parser = argparse.ArgumentParser(
            description='MCP-Sheet-Parser - Excel/CSV转HTML工具',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
示例用法:
  python main.py input.xlsx                    # 基本转换
  python main.py input.xlsx -o output.html    # 指定输出文件
  python main.py input.xlsx --theme dark      # 使用暗色主题
  python main.py input.xlsx --table-only      # 只输出表格
  python main.py input.xlsx --benchmark       # 性能测试
  python main.py input.xlsx --use-css-classes # 生成CSS类
  python main.py input.xlsx --template business # 使用商务模板
            """
        )
        
        # 基本参数
        parser.add_argument('input_file', help='输入文件路径 (Excel/CSV)')
        parser.add_argument('-o', '--output', help='输出HTML文件路径 (默认: 输入文件名.html)')
        parser.add_argument('--encoding', default='utf-8', help='文件编码 (默认: utf-8)')
        parser.add_argument('--theme', choices=['default', 'minimal', 'dark', 'print'], 
                           default='default', help='HTML主题 (默认: default)')
        parser.add_argument('--table-only', action='store_true', help='只输出表格HTML，不包含完整文档结构')
        
        # 性能参数
        performance_group = parser.add_argument_group('性能选项')
        performance_group.add_argument('--chunk-size', type=int, default=1000, 
                           help='大文件分块大小 (默认: 1000)')
        performance_group.add_argument('--max-memory', type=int, default=2048, 
                           help='最大内存使用MB (默认: 2048)')
        performance_group.add_argument('--performance-mode', choices=['auto', 'fast', 'memory'], 
                           default='auto', help='性能模式 (默认: auto)')
        performance_group.add_argument('--disable-progress', action='store_true', 
                           help='禁用进度显示')
        performance_group.add_argument('--benchmark', action='store_true', 
                           help='执行性能基准测试')
        
        # 高级样式控制参数
        style_group = parser.add_argument_group('样式选项')
        style_group.add_argument('--use-css-classes', action='store_true', 
                           help='生成CSS类而非内联样式')
        style_group.add_argument('--semantic-names', action='store_true', default=True,
                           help='使用语义化CSS类名 (默认启用)')
        style_group.add_argument('--min-usage-threshold', type=int, default=2,
                           help='CSS类最小使用次数阈值 (默认: 2)')
        style_group.add_argument('--template', choices=['business', 'financial', 'analytics'],
                           help='样式模板')
        style_group.add_argument('--conditional-rules', 
                           choices=['financial', 'analytics', 'performance', 'custom'],
                           help='预定义条件格式化规则')
        style_group.add_argument('--disable-conditional', action='store_true',
                           help='禁用条件格式化')
        
        # 公式处理参数
        formula_group = parser.add_argument_group('公式处理选项')
        formula_group.add_argument('--disable-formulas', action='store_true',
                           help='禁用公式处理')
        formula_group.add_argument('--show-formula-text', action='store_true', default=True,
                           help='在悬停时显示原始公式文本 (默认启用)')
        formula_group.add_argument('--calculate-formulas', action='store_true', default=True,
                           help='计算公式结果 (默认启用)')
        formula_group.add_argument('--show-formula-errors', action='store_true', default=True,
                           help='显示公式错误 (默认启用)')
        formula_group.add_argument('--supported-functions-only', action='store_true',
                           help='仅处理支持的函数，忽略不支持的复杂公式')
        
        # 图表转换参数
        chart_group = parser.add_argument_group('图表转换选项')
        chart_group.add_argument('--disable-charts', action='store_true',
                           help='禁用图表转换功能')
        chart_group.add_argument('--chart-format', choices=['svg', 'png'], default='svg',
                           help='图表输出格式 (默认: svg)')
        chart_group.add_argument('--chart-width', type=int, default=600,
                           help='图表默认宽度 (默认: 600)')
        chart_group.add_argument('--chart-height', type=int, default=400,
                           help='图表默认高度 (默认: 400)')
        chart_group.add_argument('--chart-quality', choices=['low', 'medium', 'high'], default='high',
                           help='图表质量 (默认: high)')
        chart_group.add_argument('--chart-responsive', action='store_true', default=True,
                           help='生成响应式图表 (默认启用)')
        
        # 调试和信息参数
        info_group = parser.add_argument_group('信息选项')
        info_group.add_argument('--verbose', '-v', action='store_true', help='详细输出')
        info_group.add_argument('--quiet', '-q', action='store_true', help='静默模式')
        info_group.add_argument('--version', action='version', version='MCP-Sheet-Parser 2.3.0')
        
        return parser
    
    def apply_config_from_args(self, args: argparse.Namespace) -> Config:
        """根据命令行参数创建配置对象"""
        config = Config()
        
        # 应用性能配置
        config.CHUNK_SIZE = args.chunk_size
        config.MAX_MEMORY_MB = args.max_memory
        config.ENABLE_PROGRESS_TRACKING = not args.disable_progress
        
        # 应用公式处理配置
        config.ENABLE_FORMULA_PROCESSING = not args.disable_formulas
        config.SHOW_FORMULA_TEXT = args.show_formula_text
        config.CALCULATE_FORMULAS = args.calculate_formulas
        config.SHOW_FORMULA_ERRORS = args.show_formula_errors
        config.SUPPORTED_FUNCTIONS_ONLY = args.supported_functions_only
        
        # 应用图表转换配置
        config.ENABLE_CHART_CONVERSION = not args.disable_charts
        config.CHART_OUTPUT_FORMAT = args.chart_format
        config.CHART_DEFAULT_WIDTH = args.chart_width
        config.CHART_DEFAULT_HEIGHT = args.chart_height
        config.CHART_QUALITY = args.chart_quality
        config.CHART_RESPONSIVE = args.chart_responsive
        
        # 根据性能模式调整配置
        if args.performance_mode == 'fast':
            config.ENABLE_PARALLEL_PROCESSING = True
            config.CHUNK_SIZE = min(args.chunk_size, 500)
        elif args.performance_mode == 'memory':
            config.ENABLE_PARALLEL_PROCESSING = False
            config.CHUNK_SIZE = max(args.chunk_size, 2000)
        
        return config
    
    def get_predefined_conditional_rules(self, rule_type: str) -> List[ConditionalRule]:
        """获取预定义条件格式化规则"""
        rules = []
        
        if rule_type == 'financial':
            rules.extend([
                ConditionalRule(
                    name="财务正值",
                    type=ConditionalType.VALUE_RANGE,
                    operator=ComparisonOperator.GREATER_THAN,
                    values=[0],
                    styles={'color': '#28a745', 'font-weight': 'bold'},
                    priority=10,
                    description="突出显示正数值"
                ),
                ConditionalRule(
                    name="财务负值",
                    type=ConditionalType.VALUE_RANGE,
                    operator=ComparisonOperator.LESS_THAN,
                    values=[0],
                    styles={'color': '#dc3545', 'font-weight': 'bold'},
                    priority=10,
                    description="突出显示负数值"
                )
            ])
        
        elif rule_type == 'analytics':
            rules.extend([
                ConditionalRule(
                    name="数据条显示",
                    type=ConditionalType.DATA_BARS,
                    operator=ComparisonOperator.GREATER_THAN,
                    values=['#4472C4'],
                    styles={},
                    priority=5,
                    description="数据条可视化"
                ),
                ConditionalRule(
                    name="重复值标记",
                    type=ConditionalType.DUPLICATE_VALUES,
                    operator=ComparisonOperator.EQUAL,
                    values=[],
                    styles={'background-color': '#fff3cd', 'color': '#856404'},
                    priority=15,
                    description="标记重复值"
                )
            ])
        
        elif rule_type == 'performance':
            rules.extend([
                ConditionalRule(
                    name="前10%",
                    type=ConditionalType.TOP_BOTTOM,
                    operator=ComparisonOperator.GREATER_THAN,
                    values=[0.1],
                    styles={'background-color': '#d4edda', 'color': '#155724', 'font-weight': 'bold'},
                    priority=20,
                    description="前10%的值"
                ),
                ConditionalRule(
                    name="后10%",
                    type=ConditionalType.TOP_BOTTOM,
                    operator=ComparisonOperator.LESS_THAN,
                    values=[0.1],
                    styles={'background-color': '#f8d7da', 'color': '#721c24', 'font-weight': 'bold'},
                    priority=20,
                    description="后10%的值"
                )
            ])
        
        return rules
    
    def setup_logging(self, args: argparse.Namespace) -> None:
        """根据参数设置日志级别"""
        import logging
        
        if args.quiet:
            log_level = logging.WARNING
        elif args.verbose:
            log_level = logging.DEBUG
        else:
            log_level = logging.INFO
        
        logging.basicConfig(
            level=log_level,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )


# 创建全局CLI管理器实例
cli_manager = CLIManager()

# 导出主要函数供main.py使用
create_argument_parser = cli_manager.create_argument_parser
apply_config_from_args = cli_manager.apply_config_from_args
get_predefined_conditional_rules = cli_manager.get_predefined_conditional_rules
setup_logging = cli_manager.setup_logging 