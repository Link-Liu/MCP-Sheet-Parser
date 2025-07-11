#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# main.py
# MCP-Sheet-Parser 主程序入口

import os
import sys
from typing import Optional
from mcp_sheet_parser.cli import CLIManager
from mcp_sheet_parser.file_processor import FileProcessor
from mcp_sheet_parser.exceptions import (
    ErrorHandler, ErrorSeverity, MCPSheetParserError,
    create_error_context, global_error_handler
)


def create_progress_callback():
    """创建进度回调函数"""
    def progress_callback(progress: int, message: str):
        # 简单的进度显示
        bar_length = 30
        filled_length = int(bar_length * progress // 100)
        bar = '█' * filled_length + '-' * (bar_length - filled_length)
        print(f'\r进度: |{bar}| {progress}% {message}', end='', flush=True)
        if progress >= 100:
            print()  # 完成时换行
    return progress_callback


def handle_global_error(error: Exception) -> int:
    """处理全局错误并返回退出代码"""
    context = create_error_context("程序执行")
    
    # 使用全局错误处理器
    if isinstance(error, MCPSheetParserError):
        custom_error = error
    else:
        custom_error = global_error_handler.handle_error(error, context)
    
    # 根据错误严重程度确定退出代码
    if custom_error.severity == ErrorSeverity.CRITICAL:
        exit_code = 3
    elif custom_error.severity == ErrorSeverity.HIGH:
        exit_code = 2
    else:
        exit_code = 1
    
    print(f"\n程序执行失败: {custom_error.get_detailed_message()}", file=sys.stderr)
    
    # 显示错误统计
    error_summary = global_error_handler.get_error_summary()
    if error_summary['total_errors'] > 1:
        print(f"总错误数: {error_summary['total_errors']}", file=sys.stderr)
        print("错误类型分布:", file=sys.stderr)
        for error_type, count in error_summary['by_type'].items():
            print(f"  {error_type}: {count}", file=sys.stderr)
    
    return exit_code


def main():
    """主函数"""
    try:
        # 初始化CLI处理器
        cli_manager = CLIManager()
        
        # 解析命令行参数
        parser = cli_manager.create_argument_parser()
        args = parser.parse_args()
        
        # 创建配置
        config = cli_manager.apply_config_from_args(args)
        
        # 设置日志
        cli_manager.setup_logging(args)
        
        print("MCP-Sheet-Parser - 表格文件转HTML工具")
        print("=" * 50)
        
        # 创建文件处理器
        processor = FileProcessor(config)
        
        # 验证输入文件
        validation = processor.validate_input_file(args.input_file)
        if not validation['is_valid']:
            print("输入文件验证失败:", file=sys.stderr)
            for error in validation['errors']:
                print(f"  错误: {error}", file=sys.stderr)
            return 2
        
        # 确定输出文件路径
        output_file = args.output
        if not output_file:
            # 如果没有指定输出文件，生成默认文件名
            base_name = os.path.splitext(os.path.basename(args.input_file))[0]
            output_file = f"{base_name}.html"
        
        # 确定页面标题
        title = getattr(args, 'title', None)
        if not title:
            title = os.path.splitext(os.path.basename(args.input_file))[0]
        
        # 显示文件信息
        print(f"输入文件: {args.input_file}")
        print(f"文件大小: {validation['file_size_mb']} MB")
        print(f"输出文件: {output_file}")
        print(f"主题: {args.theme}")
        
        # 显示警告
        if validation['warnings']:
            for warning in validation['warnings']:
                print(f"警告: {warning}")
        
        print("\n开始处理...")
        
        # 创建进度回调
        progress_callback = create_progress_callback()
        
        # 处理文件
        result = processor.process_file(
            input_path=args.input_file,
            output_path=output_file,
            theme=args.theme,
            title=title,
            progress_callback=progress_callback
        )
        
        if result['success']:
            print(f"\n处理完成!")
            print(f"工作表数量: {result['sheets_count']}")
            print(f"处理时间: {result['processing_time']:.2f} 秒")
            print(f"输出文件: {result['output_path']}")
            
            # 显示统计信息
            stats = processor.get_processing_stats()
            if stats['error_summary']['total_errors'] > 0:
                print(f"\n处理过程中的警告/错误: {stats['error_summary']['total_errors']}")
                print("详细信息:")
                for error_type, count in stats['error_summary']['by_type'].items():
                    print(f"  {error_type}: {count}")
            
            return 0
        else:
            print(f"\n处理失败: {result['error_message']}", file=sys.stderr)
            return 1
    
    except KeyboardInterrupt:
        print("\n\n用户中断操作", file=sys.stderr)
        return 130  # SIGINT exit code
    
    except Exception as e:
        return handle_global_error(e)


if __name__ == "__main__":
    sys.exit(main()) 