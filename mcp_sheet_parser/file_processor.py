# file_processor.py
# 文件处理核心逻辑 - 处理文件解析和转换的完整流程

import os
import time
from typing import List, Dict, Any, Optional, Callable
from .parser import SheetParser
from .html_converter import HTMLConverter
from .config import Config
from .utils import setup_logger, get_file_extension
from .exceptions import (
    ErrorHandler, ErrorContext, ErrorSeverity, MCPSheetParserError,
    FileProcessingError, ParsingError, ConversionError, 
    error_handler, safe_execute, create_error_context
)


class FileProcessor:
    """文件处理器 - 统一的文件处理接口"""
    
    def __init__(self, config: Optional[Config] = None):
        self.config = config or Config()
        self.logger = setup_logger(__name__)
        self.error_handler = ErrorHandler(self.logger)
        
        # 处理统计
        self.stats = {
            'total_files': 0,
            'successful_files': 0,
            'failed_files': 0,
            'start_time': None,
            'end_time': None,
            'processing_time': 0
        }
    
    @error_handler(operation="文件处理")
    def process_file(self, 
                    input_path: str, 
                    output_path: str, 
                    theme: str = 'default',
                    title: Optional[str] = None,
                    progress_callback: Optional[Callable] = None) -> Dict[str, Any]:
        """
        处理单个文件的完整流程
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            theme: HTML主题
            title: HTML页面标题
            progress_callback: 进度回调函数
            
        Returns:
            Dict: 处理结果信息
        """
        context = create_error_context(
            "文件处理",
            file_path=input_path,
            additional_info={
                'output_path': output_path,
                'theme': theme,
                'title': title
            }
        )
        
        self.stats['start_time'] = time.time()
        self.stats['total_files'] += 1
        
        try:
            # 生成默认标题
            if title is None:
                filename = os.path.basename(input_path)
                title = os.path.splitext(filename)[0]
            
            self.logger.info(f"开始处理文件: {input_path}")
            
            # 阶段1: 解析文件
            if progress_callback:
                progress_callback(10, "正在解析文件...")
            
            sheets_data = self._parse_file(input_path, progress_callback, context)
            
            if not sheets_data:
                raise ParsingError("文件解析失败或文件为空", context=context)
            
            # 阶段2: 转换为HTML
            if progress_callback:
                progress_callback(80, "正在生成HTML...")
            
            result_path = self._convert_to_html(
                sheets_data, output_path, theme, title, context
            )
            
            # 阶段3: 完成
            if progress_callback:
                progress_callback(100, "处理完成")
            
            self.stats['successful_files'] += 1
            processing_time = time.time() - self.stats['start_time']
            self.stats['processing_time'] += processing_time
            
            result = {
                'success': True,
                'input_path': input_path,
                'output_path': result_path,
                'sheets_count': len(sheets_data),
                'processing_time': processing_time,
                'theme': theme,
                'title': title
            }
            
            self.logger.info(f"文件处理完成: {input_path} -> {result_path}")
            return result
            
        except Exception as e:
            self.stats['failed_files'] += 1
            
            # 使用统一错误处理器
            if isinstance(e, MCPSheetParserError):
                custom_error = e
            else:
                custom_error = self.error_handler.handle_error(e, context)
            
            error_result = {
                'success': False,
                'input_path': input_path,
                'error_type': type(custom_error).__name__,
                'error_message': custom_error.get_detailed_message(),
                'error_severity': custom_error.severity.value,
                'processing_time': time.time() - self.stats['start_time'] if self.stats['start_time'] else 0
            }
            
            self.logger.error(f"文件处理失败: {input_path} - {custom_error.get_detailed_message()}")
            
            if progress_callback:
                progress_callback(100, f"处理失败: {custom_error.message}")
            
            # 根据错误严重程度决定是否重新抛出
            if custom_error.severity in [ErrorSeverity.HIGH, ErrorSeverity.CRITICAL]:
                raise custom_error
            
            return error_result
        
        finally:
            self.stats['end_time'] = time.time()
    
    def _parse_file(self, input_path: str, progress_callback: Optional[Callable], 
                   context: ErrorContext) -> List[Dict]:
        """解析文件"""
        parse_context = create_error_context(
            "文件解析",
            file_path=input_path
        )
        
        try:
            # 创建解析器
            parser = SheetParser(
                file_path=input_path,
                config=self.config,
                progress_callback=progress_callback
            )
            
            # 执行解析
            sheets_data = parser.parse()
            
            if progress_callback:
                progress_callback(60, f"解析完成，共{len(sheets_data)}个工作表")
            
            self.logger.info(f"文件解析成功: {len(sheets_data)}个工作表")
            return sheets_data
            
        except Exception as e:
            if isinstance(e, MCPSheetParserError):
                raise e
            else:
                custom_error = self.error_handler.handle_error(e, parse_context)
                raise custom_error
    
    def _convert_to_html(self, sheets_data: List[Dict], output_path: str,
                        theme: str, title: str, context: ErrorContext) -> str:
        """转换为HTML"""
        convert_context = create_error_context(
            "HTML转换",
            file_path=output_path,
            additional_info={'theme': theme, 'sheets_count': len(sheets_data)}
        )
        
        try:
            # 创建HTML转换器
            converter = HTMLConverter(config=self.config)
            
            # 执行转换
            result_path = converter.convert_to_html(
                sheets_data=sheets_data,
                output_path=output_path,
                theme=theme,
                title=title,
                include_styles=True
            )
            
            self.logger.info(f"HTML转换成功: {result_path}")
            return result_path
            
        except Exception as e:
            if isinstance(e, MCPSheetParserError):
                raise e
            else:
                custom_error = self.error_handler.handle_error(e, convert_context)
                raise custom_error
    
    def process_multiple_files(self, 
                             file_paths: List[str],
                             output_dir: str,
                             theme: str = 'default',
                             progress_callback: Optional[Callable] = None) -> List[Dict[str, Any]]:
        """
        批量处理多个文件
        
        Args:
            file_paths: 输入文件路径列表
            output_dir: 输出目录
            theme: HTML主题
            progress_callback: 进度回调函数
            
        Returns:
            List[Dict]: 每个文件的处理结果
        """
        context = create_error_context(
            "批量文件处理",
            additional_info={
                'file_count': len(file_paths),
                'output_dir': output_dir,
                'theme': theme
            }
        )
        
        try:
            self.logger.info(f"开始批量处理 {len(file_paths)} 个文件")
            
            # 确保输出目录存在
            os.makedirs(output_dir, exist_ok=True)
            
            results = []
            total_files = len(file_paths)
            
            for i, input_path in enumerate(file_paths):
                try:
                    # 生成输出文件名
                    base_name = os.path.splitext(os.path.basename(input_path))[0]
                    output_path = os.path.join(output_dir, f"{base_name}.html")
                    
                    # 更新整体进度
                    if progress_callback:
                        overall_progress = int((i / total_files) * 100)
                        progress_callback(overall_progress, f"处理文件 {i+1}/{total_files}: {os.path.basename(input_path)}")
                    
                    # 创建单个文件的进度回调
                    def file_progress(progress, message):
                        if progress_callback:
                            # 计算总体进度
                            file_weight = 100 / total_files
                            total_progress = int((i / total_files) * 100 + (progress / 100) * file_weight)
                            progress_callback(total_progress, f"文件 {i+1}/{total_files}: {message}")
                    
                    # 处理单个文件
                    result = self.process_file(
                        input_path=input_path,
                        output_path=output_path,
                        theme=theme,
                        progress_callback=file_progress
                    )
                    
                    results.append(result)
                    
                except Exception as e:
                    # 单个文件失败不影响其他文件
                    file_error = safe_execute(
                        self.error_handler.handle_error,
                        e,
                        create_error_context("单文件处理", file_path=input_path),
                        operation=f"单文件处理错误-{input_path}",
                        default_value=MCPSheetParserError(f"文件处理失败: {e}")
                    )
                    
                    error_result = {
                        'success': False,
                        'input_path': input_path,
                        'error_type': type(file_error).__name__,
                        'error_message': file_error.get_detailed_message() if hasattr(file_error, 'get_detailed_message') else str(file_error),
                        'error_severity': file_error.severity.value if hasattr(file_error, 'severity') else 'medium'
                    }
                    
                    results.append(error_result)
                    self.logger.warning(f"跳过失败文件: {input_path}")
                    continue
            
            # 最终进度更新
            if progress_callback:
                progress_callback(100, f"批量处理完成: {self.stats['successful_files']}/{total_files} 成功")
            
            self.logger.info(f"批量处理完成: {self.stats['successful_files']}/{total_files} 成功")
            return results
            
        except Exception as e:
            custom_error = self.error_handler.handle_error(e, context)
            raise custom_error
    
    def get_processing_stats(self) -> Dict[str, Any]:
        """获取处理统计信息"""
        stats = self.stats.copy()
        
        # 计算成功率
        if stats['total_files'] > 0:
            stats['success_rate'] = (stats['successful_files'] / stats['total_files']) * 100
        else:
            stats['success_rate'] = 0
        
        # 计算平均处理时间
        if stats['successful_files'] > 0:
            stats['avg_processing_time'] = stats['processing_time'] / stats['successful_files']
        else:
            stats['avg_processing_time'] = 0
        
        # 添加错误统计
        stats['error_summary'] = self.error_handler.get_error_summary()
        
        return stats
    
    def reset_stats(self):
        """重置统计信息"""
        self.stats = {
            'total_files': 0,
            'successful_files': 0,
            'failed_files': 0,
            'start_time': None,
            'end_time': None,
            'processing_time': 0
        }
        self.error_handler.reset_stats()
    
    def get_supported_formats(self) -> List[str]:
        """获取支持的文件格式"""
        return list(self.config.get_all_supported_formats())
    
    def validate_input_file(self, file_path: str) -> Dict[str, Any]:
        """验证输入文件"""
        context = create_error_context("文件验证", file_path=file_path)
        
        validation_result = {
            'is_valid': True,
            'file_path': file_path,
            'file_exists': False,
            'is_supported_format': False,
            'file_size_mb': 0,
            'warnings': [],
            'errors': []
        }
        
        try:
            # 检查文件是否存在
            if not os.path.exists(file_path):
                validation_result['is_valid'] = False
                validation_result['errors'].append("文件不存在")
                return validation_result
            
            validation_result['file_exists'] = True
            
            # 检查文件格式
            ext = get_file_extension(file_path)
            supported_formats = self.get_supported_formats()
            
            if ext not in supported_formats:
                validation_result['is_valid'] = False
                validation_result['errors'].append(f"不支持的文件格式: {ext}")
            else:
                validation_result['is_supported_format'] = True
            
            # 检查文件大小
            file_size_bytes = os.path.getsize(file_path)
            file_size_mb = file_size_bytes / (1024 * 1024)
            validation_result['file_size_mb'] = round(file_size_mb, 2)
            
            if file_size_mb > self.config.MAX_FILE_SIZE_MB:
                validation_result['is_valid'] = False
                validation_result['errors'].append(
                    f"文件过大: {file_size_mb:.2f}MB > {self.config.MAX_FILE_SIZE_MB}MB"
                )
            elif file_size_mb > 100:
                validation_result['warnings'].append(
                    f"大文件警告: {file_size_mb:.2f}MB，处理可能较慢"
                )
            
            return validation_result
            
        except Exception as e:
            custom_error = self.error_handler.handle_error(e, context)
            validation_result['is_valid'] = False
            validation_result['errors'].append(f"验证失败: {custom_error.message}")
            return validation_result 