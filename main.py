# main.py
# MCP-Sheet-Parser 程序入口

import sys
import argparse
import os
import glob
from pathlib import Path
import time # Added for performance benchmark

from mcp_sheet_parser.parser import SheetParser
from mcp_sheet_parser.html_converter import HTMLConverter
from mcp_sheet_parser.config import Config, THEMES
from mcp_sheet_parser.utils import (
    setup_logger, 
    batch_process_files, 
    get_file_info, 
    format_file_size,
    is_supported_format
)


def create_parser():
    """创建命令行参数解析器"""
    parser = argparse.ArgumentParser(
        description='MCP-Sheet-Parser: 表格文件解析器，专注于HTML转换',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  python main.py input.xlsx --html output.html
  python main.py input.csv --html output.html --theme dark
  python main.py *.xlsx --batch --output-dir ./html_files/
  python main.py input.xlsx --info
        """
    )
    
    # 必需参数
    parser.add_argument(
        'input',
        nargs='*',
        help='输入文件路径（支持通配符和多个文件）'
    )
    
    # 输出选项
    output_group = parser.add_argument_group('输出选项')
    output_group.add_argument(
        '--html', '-o',
        metavar='PATH',
        help='输出HTML文件路径'
    )
    output_group.add_argument(
        '--output-dir', '-d',
        metavar='DIR',
        help='批量处理时的输出目录'
    )
    output_group.add_argument(
        '--table-only',
        action='store_true',
        help='只输出表格HTML，不包含完整文档结构'
    )
    
    # 样式选项
    style_group = parser.add_argument_group('样式选项')
    style_group.add_argument(
        '--theme', '-t',
        choices=list(THEMES.keys()),
        default='default',
        help='HTML输出主题 (默认: default)'
    )
    style_group.add_argument(
        '--no-comments',
        action='store_true',
        help='不包含单元格批注'
    )
    style_group.add_argument(
        '--no-hyperlinks',
        action='store_true',
        help='不包含超链接'
    )
    
    # 处理选项
    process_group = parser.add_argument_group('处理选项')
    process_group.add_argument(
        '--batch', '-b',
        action='store_true',
        help='批量处理模式'
    )
    process_group.add_argument(
        '--sheet',
        type=int,
        metavar='N',
        help='只处理指定工作表（从0开始）'
    )
    process_group.add_argument(
        '--chunk-size',
        type=int,
        default=1000,
        metavar='N',
        help='分块处理大小（行数，默认1000）'
    )
    process_group.add_argument(
        '--max-memory',
        type=int,
        default=2048,
        metavar='MB',
        help='最大内存使用（MB，默认2048）'
    )
    process_group.add_argument(
        '--disable-progress',
        action='store_true',
        help='禁用进度跟踪'
    )
    process_group.add_argument(
        '--performance-mode',
        action='store_true',
        help='启用高性能模式（大文件优化）'
    )
    
    # 信息选项
    info_group = parser.add_argument_group('信息选项')
    info_group.add_argument(
        '--info', '-i',
        action='store_true',
        help='显示文件信息'
    )
    info_group.add_argument(
        '--list-themes',
        action='store_true',
        help='显示可用主题列表'
    )
    info_group.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='详细输出'
    )
    info_group.add_argument(
        '--version',
        action='version',
        version='MCP-Sheet-Parser 1.0.0'
    )
    
    return parser


def list_themes():
    """显示主题列表"""
    print("\n🎨 可用主题:")
    for theme_name, theme_config in THEMES.items():
        print(f"  {theme_name}: {theme_config['description']}")
    print()


def show_file_info(file_path):
    """显示文件信息"""
    info = get_file_info(file_path)
    
    if 'error' in info:
        print(f"❌ {info['error']}")
        return
    
    print(f"\n📄 文件信息: {os.path.basename(file_path)}")
    print(f"  路径: {info['path']}")
    print(f"  大小: {info['size_formatted']}")
    print(f"  扩展名: {info['extension']}")
    print(f"  支持: {'✅ 是' if info['is_supported'] else '❌ 否'}")


def setup_config(args):
    """根据命令行参数设置配置"""
    config = Config()
    
    if args.no_comments:
        config.INCLUDE_COMMENTS = False
    
    if args.no_hyperlinks:
        config.INCLUDE_HYPERLINKS = False
    
    # 性能配置
    if hasattr(args, 'chunk_size'):
        config.CHUNK_SIZE = args.chunk_size
    
    if hasattr(args, 'max_memory'):
        config.MAX_MEMORY_MB = args.max_memory
    
    if hasattr(args, 'disable_progress'):
        config.ENABLE_PROGRESS_TRACKING = not args.disable_progress
    
    if hasattr(args, 'performance_mode') and args.performance_mode:
        config.ENABLE_PERFORMANCE_MODE = True
        # 高性能模式下的优化设置
        config.CHUNK_SIZE = min(args.chunk_size, 500)  # 减小块大小
        config.MAX_MEMORY_MB = max(args.max_memory, 4096)  # 增加内存限制
    
    return config


def generate_output_path(input_path, output_dir, theme):
    """生成输出文件路径"""
    base_name = Path(input_path).stem
    if theme != 'default':
        output_name = f"{base_name}_{theme}.html"
    else:
        output_name = f"{base_name}.html"
    
    return os.path.join(output_dir, output_name)


def progress_callback(current, total, current_file):
    """进度回调函数"""
    percentage = (current / total) * 100
    print(f"\r进度: {current}/{total} ({percentage:.1f}%) - {os.path.basename(current_file)}", end='')
    if current == total:
        print()  # 完成时换行


def process_single_file(input_path, args, config):
    """处理单个文件"""
    logger = setup_logger(__name__)
    
    try:
        # 创建进度回调
        progress_callback = None
        if (getattr(config, 'ENABLE_PROGRESS_TRACKING', True) and 
            not getattr(args, 'disable_progress', False) and 
            args.verbose):
            from mcp_sheet_parser.performance import create_progress_callback
            progress_callback = create_progress_callback(verbose=True)
        
        # 解析文件
        parser = SheetParser(input_path, config, progress_callback)
        sheets = parser.parse()
        
        if not sheets:
            print(f"⚠️  文件解析结果为空: {input_path}")
            return False
        
        # 显示基本信息
        if args.verbose:
            print(f"\n📊 解析完成: {os.path.basename(input_path)}")
            for i, sheet in enumerate(sheets):
                print(f"  工作表 {i}: {sheet['sheet_name']} ({sheet['rows']}行 × {sheet['cols']}列)")
            if sheet['merged_cells']:
                    print(f"    合并单元格: {len(sheet['merged_cells'])}个")
        
        # 如果只显示信息，则返回
        if args.info:
            return True
        
        # 确定要处理的工作表
        if args.sheet is not None:
            if 0 <= args.sheet < len(sheets):
                sheets = [sheets[args.sheet]]
            else:
                print(f"❌ 工作表序号超出范围: {args.sheet} (总共 {len(sheets)} 个工作表)")
                return False
        
        # 处理每个工作表
        success = True
        for i, sheet_data in enumerate(sheets):
            try:
                # 创建HTML转换器
                converter = HTMLConverter(sheet_data, config, args.theme)
                
                # 确定输出路径
                if args.html:
                    output_path = args.html
                    if len(sheets) > 1:
                        # 多个工作表时添加索引
                        base, ext = os.path.splitext(args.html)
                        output_path = f"{base}_sheet{i}{ext}"
                elif args.output_dir:
                    output_path = generate_output_path(input_path, args.output_dir, args.theme)
                    if len(sheets) > 1:
                        base, ext = os.path.splitext(output_path)
                        output_path = f"{base}_sheet{i}{ext}"
                else:
                    # 默认输出路径
                    base_name = Path(input_path).stem
                    output_path = f"{base_name}.html"
                
                # 导出HTML
                converter.export_to_file(output_path, args.table_only)
                
                if args.verbose:
                    print(f"✅ HTML已导出: {output_path}")
                
            except Exception as e:
                logger.error(f"处理工作表失败 {sheet_data['sheet_name']}: {e}")
                success = False
        
        return success
        
    except Exception as e:
        logger.error(f"处理文件失败 {input_path}: {e}")
        return False


def main():
    """主函数"""
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
  python main.py input.xlsx --conditional-rules financial # 应用财务条件格式化
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
    parser.add_argument('--chunk-size', type=int, default=1000, 
                       help='大文件分块大小 (默认: 1000)')
    parser.add_argument('--max-memory', type=int, default=2048, 
                       help='最大内存使用MB (默认: 2048)')
    parser.add_argument('--performance-mode', choices=['auto', 'fast', 'memory'], 
                       default='auto', help='性能模式 (默认: auto)')
    parser.add_argument('--disable-progress', action='store_true', 
                       help='禁用进度显示')
    parser.add_argument('--benchmark', action='store_true', 
                       help='执行性能基准测试')
    
    # 高级样式控制参数
    parser.add_argument('--use-css-classes', action='store_true', 
                       help='生成CSS类而非内联样式')
    parser.add_argument('--semantic-names', action='store_true', default=True,
                       help='使用语义化CSS类名 (默认启用)')
    parser.add_argument('--min-usage-threshold', type=int, default=2,
                       help='CSS类最小使用次数阈值 (默认: 2)')
    parser.add_argument('--template', choices=['business', 'financial', 'analytics'],
                       help='样式模板')
    parser.add_argument('--conditional-rules', 
                       choices=['financial', 'analytics', 'performance', 'custom'],
                       help='预定义条件格式化规则')
    parser.add_argument('--disable-conditional', action='store_true',
                       help='禁用条件格式化')
    
    # 公式处理参数
    parser.add_argument('--disable-formulas', action='store_true',
                       help='禁用公式处理')
    parser.add_argument('--show-formula-text', action='store_true', default=True,
                       help='在悬停时显示原始公式文本 (默认启用)')
    parser.add_argument('--calculate-formulas', action='store_true', default=True,
                       help='计算公式结果 (默认启用)')
    parser.add_argument('--show-formula-errors', action='store_true', default=True,
                       help='显示公式错误 (默认启用)')
    parser.add_argument('--supported-functions-only', action='store_true',
                       help='仅处理支持的函数，忽略不支持的复杂公式')
    
    # 图表转换参数
    parser.add_argument('--disable-charts', action='store_true',
                       help='禁用图表转换功能')
    parser.add_argument('--chart-format', choices=['svg', 'png'], default='svg',
                       help='图表输出格式 (默认: svg)')
    parser.add_argument('--chart-width', type=int, default=600,
                       help='图表默认宽度 (默认: 600)')
    parser.add_argument('--chart-height', type=int, default=400,
                       help='图表默认高度 (默认: 400)')
    parser.add_argument('--chart-quality', choices=['low', 'medium', 'high'], default='high',
                       help='图表质量 (默认: high)')
    parser.add_argument('--chart-responsive', action='store_true', default=True,
                       help='生成响应式图表 (默认启用)')
    
    # 调试和信息参数
    parser.add_argument('--verbose', '-v', action='store_true', help='详细输出')
    parser.add_argument('--quiet', '-q', action='store_true', help='静默模式')
    parser.add_argument('--version', action='version', version='MCP-Sheet-Parser 2.3.0')
    
    args = parser.parse_args()
    
    # 设置日志级别
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
    
    try:
        # 验证输入文件
        if not os.path.exists(args.input_file):
            print(f"❌ 错误: 输入文件不存在: {args.input_file}")
            return 1
        
        # 生成输出文件名
        if args.output:
            output_file = args.output
        else:
            base_name = os.path.splitext(args.input_file)[0]
            output_file = f"{base_name}.html"
        
        print(f"🔄 正在处理: {args.input_file}")
        print(f"📄 输出文件: {output_file}")
        print(f"🎨 使用主题: {args.theme}")
        
        # 性能基准测试
        if args.benchmark:
            print("🚀 执行性能基准测试...")
            from mcp_sheet_parser.performance import PerformanceBenchmark
            
            benchmark = PerformanceBenchmark()
            results = benchmark.run_comprehensive_benchmark(args.input_file)
            
            print("\n📊 基准测试结果:")
            for metric, value in results.items():
                print(f"  {metric}: {value}")
            return 0
        
        # 创建配置
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
        
        # 解析文件
        print("📊 开始解析文件...")
        start_time = time.time()
        
        # 创建进度回调函数
        def create_progress_callback():
            if args.disable_progress:
                return None
            
            def callback(progress_info):
                percent = progress_info.get('percentage', 0)
                processed = progress_info.get('processed', 0)
                total = progress_info.get('total', 1)
                message = progress_info.get('description', '')
                print(f"\r进度: {percent:.1f}% ({processed}/{total}) {message}", end='', flush=True)
            return callback
        
        # 创建解析器
        parser = SheetParser(args.input_file, config, create_progress_callback())
        
        # 解析数据
        sheet_data_list = parser.parse()
        
        # 取第一个工作表数据
        if sheet_data_list:
            sheet_data = sheet_data_list[0]
        else:
            raise ValueError("文件解析失败或文件为空")
        
        parse_time = time.time() - start_time
        print(f"\n✅ 解析完成，用时: {parse_time:.2f}秒")
        
        # 准备条件格式化规则
        conditional_rules = None
        if args.conditional_rules and not args.disable_conditional:
            conditional_rules = get_predefined_conditional_rules(args.conditional_rules)
        
        # 创建HTML转换器
        print("🌐 开始HTML转换...")
        html_start_time = time.time()
        
        converter = HTMLConverter(
            sheet_data, 
            config=config, 
            theme=args.theme,
            use_css_classes=args.use_css_classes,
            conditional_rules=conditional_rules
        )
        
        # 设置样式模板
        if args.template:
            converter.set_style_template(args.template)
            print(f"🎨 应用样式模板: {args.template}")
        
        # 配置样式选项
        style_options = {
            'use_css_classes': args.use_css_classes,
            'semantic_names': args.semantic_names,
            'min_usage_threshold': args.min_usage_threshold,
            'apply_conditional': not args.disable_conditional,
            'template': args.template
        }
        
        # 生成HTML
        html_content = converter.to_html(
            table_only=args.table_only,
            style_options=style_options
        )
        
        html_time = time.time() - html_start_time
        print(f"✅ HTML转换完成，用时: {html_time:.2f}秒")
        
        # 保存文件
        print("💾 保存HTML文件...")
        write_start_time = time.time()
        
        with open(output_file, 'w', encoding=args.encoding) as f:
            f.write(html_content)
        
        write_time = time.time() - write_start_time
        total_time = time.time() - start_time
        
        print(f"✅ 文件保存完成，用时: {write_time:.2f}秒")
        print(f"🎉 总用时: {total_time:.2f}秒")
        
        # 显示样式统计信息
        if args.use_css_classes and args.verbose:
            stats = converter.get_style_statistics()
            print(f"\n📈 样式统计:")
            print(f"  总样式数: {stats.get('total_styles', 0)}")
            print(f"  独特样式数: {stats.get('unique_styles', 0)}")
            print(f"  CSS类复用率: {stats.get('class_reuse_rate', 0):.1f}%")
            print(f"  条件格式化应用: {stats.get('conditional_rules_applied', 0)}次")
        
        # 显示文件信息
        file_size = os.path.getsize(output_file)
        print(f"📄 输出文件大小: {file_size / 1024:.1f} KB")
        
        return 0
        
    except FileNotFoundError as e:
        print(f"❌ 文件错误: {e}")
        return 1
    except PermissionError as e:
        print(f"❌ 权限错误: {e}")
        return 1
    except Exception as e:
        print(f"❌ 转换失败: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


def get_predefined_conditional_rules(rule_type):
    """获取预定义条件格式化规则"""
    from mcp_sheet_parser.style_manager import ConditionalRule, ConditionalType, ComparisonOperator
    
    rules = []
    
    if rule_type == 'financial':
        # 财务规则：正值绿色，负值红色
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
        # 分析规则：颜色渐变和数据条
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
        # 性能规则：前10%和后10%
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


if __name__ == "__main__":
    import logging
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n⏹️  用户中断操作")
        sys.exit(130)
    except Exception as e:
        print(f"\n💥 程序异常: {e}")
        sys.exit(1)
