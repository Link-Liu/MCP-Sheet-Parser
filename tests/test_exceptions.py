# test_exceptions.py
# 测试新的异常处理体系

import unittest
import tempfile
import os
from unittest.mock import patch, MagicMock

from mcp_sheet_parser.exceptions import (
    ErrorHandler, ErrorContext, ErrorSeverity, MCPSheetParserError,
    FileNotFoundError, UnsupportedFormatError, FileSizeExceededError,
    WorksheetParsingError, CellProcessingError, ParsingError,
    HTMLConversionError, FormulaProcessingError, SecurityError,
    PerformanceError, ConfigurationError,
    error_handler, safe_execute, create_error_context, global_error_handler
)


class TestErrorContext(unittest.TestCase):
    """测试错误上下文"""
    
    def test_error_context_creation(self):
        """测试错误上下文创建"""
        context = ErrorContext(
            operation="测试操作",
            file_path="/test/file.xlsx",
            sheet_name="Sheet1",
            cell_position="A1",
            additional_info={'key': 'value'}
        )
        
        self.assertEqual(context.operation, "测试操作")
        self.assertEqual(context.file_path, "/test/file.xlsx")
        self.assertEqual(context.sheet_name, "Sheet1")
        self.assertEqual(context.cell_position, "A1")
        self.assertEqual(context.additional_info['key'], 'value')
    
    def test_error_context_to_dict(self):
        """测试错误上下文转字典"""
        context = ErrorContext(
            operation="测试操作",
            file_path="/test/file.xlsx"
        )
        
        context_dict = context.to_dict()
        self.assertIsInstance(context_dict, dict)
        self.assertEqual(context_dict['operation'], "测试操作")
        self.assertEqual(context_dict['file_path'], "/test/file.xlsx")
    
    def test_create_error_context_helper(self):
        """测试错误上下文创建辅助函数"""
        context = create_error_context(
            "测试操作",
            file_path="/test/file.xlsx",
            sheet_name="Sheet1"
        )
        
        self.assertEqual(context.operation, "测试操作")
        self.assertEqual(context.file_path, "/test/file.xlsx")
        self.assertEqual(context.sheet_name, "Sheet1")


class TestCustomExceptions(unittest.TestCase):
    """测试自定义异常类"""
    
    def test_base_exception(self):
        """测试基础异常类"""
        context = ErrorContext(operation="测试")
        error = MCPSheetParserError(
            "测试错误",
            severity=ErrorSeverity.HIGH,
            context=context
        )
        
        self.assertEqual(error.message, "测试错误")
        self.assertEqual(error.severity, ErrorSeverity.HIGH)
        self.assertEqual(error.context, context)
        
        # 测试详细消息
        detailed_msg = error.get_detailed_message()
        self.assertIn("HIGH", detailed_msg)
        self.assertIn("测试错误", detailed_msg)
        self.assertIn("操作: 测试", detailed_msg)
    
    def test_file_not_found_error(self):
        """测试文件未找到错误"""
        error = FileNotFoundError("/nonexistent/file.xlsx")
        
        self.assertIn("/nonexistent/file.xlsx", error.message)
        self.assertEqual(error.severity, ErrorSeverity.HIGH)
        self.assertEqual(error.context.file_path, "/nonexistent/file.xlsx")
    
    def test_unsupported_format_error(self):
        """测试不支持格式错误"""
        error = UnsupportedFormatError("/test/file.abc", ".abc")
        
        self.assertIn(".abc", error.message)
        self.assertIn("/test/file.abc", error.message)
        self.assertEqual(error.severity, ErrorSeverity.HIGH)
    
    def test_file_size_exceeded_error(self):
        """测试文件大小超限错误"""
        error = FileSizeExceededError("/test/large.xlsx", 1000.5, 500)
        
        self.assertIn("1000.5", error.message)
        self.assertIn("500", error.message)
        self.assertEqual(error.severity, ErrorSeverity.HIGH)
    
    def test_worksheet_parsing_error(self):
        """测试工作表解析错误"""
        error = WorksheetParsingError("Sheet1", "解析失败")
        
        self.assertIn("Sheet1", error.message)
        self.assertIn("解析失败", error.message)
        self.assertEqual(error.severity, ErrorSeverity.MEDIUM)
    
    def test_cell_processing_error(self):
        """测试单元格处理错误"""
        error = CellProcessingError("A1", "单元格错误")
        
        self.assertIn("A1", error.message)
        self.assertIn("单元格错误", error.message)
        self.assertEqual(error.severity, ErrorSeverity.LOW)
    
    def test_exception_to_dict(self):
        """测试异常转字典"""
        context = ErrorContext(operation="测试")
        error = MCPSheetParserError("测试", context=context)
        
        error_dict = error.to_dict()
        self.assertEqual(error_dict['error_type'], 'MCPSheetParserError')
        self.assertEqual(error_dict['message'], '测试')
        self.assertEqual(error_dict['severity'], 'medium')
        self.assertIsNotNone(error_dict['context'])


class TestErrorHandler(unittest.TestCase):
    """测试错误处理器"""
    
    def setUp(self):
        """设置测试"""
        self.error_handler = ErrorHandler()
    
    def test_handle_custom_error(self):
        """测试处理自定义错误"""
        original_error = MCPSheetParserError("测试错误")
        
        handled_error = self.error_handler.handle_error(original_error)
        
        self.assertEqual(handled_error, original_error)
        self.assertEqual(self.error_handler.get_error_summary()['total_errors'], 1)
    
    def test_handle_builtin_exception(self):
        """测试处理内置异常"""
        original_error = ValueError("值错误")
        context = ErrorContext(operation="测试")
        
        handled_error = self.error_handler.handle_error(original_error, context)
        
        self.assertIsInstance(handled_error, ParsingError)
        self.assertEqual(handled_error.original_error, original_error)
        self.assertEqual(handled_error.context, context)
    
    def test_handle_file_not_found(self):
        """测试处理文件未找到异常"""
        original_error = FileNotFoundError("文件不存在")
        
        handled_error = self.error_handler.handle_error(original_error)
        
        # 应该使用我们的自定义FileNotFoundError
        self.assertIsInstance(handled_error, FileNotFoundError)
    
    def test_handle_permission_error(self):
        """测试处理权限错误"""
        original_error = PermissionError("权限不足")
        
        handled_error = self.error_handler.handle_error(original_error)
        
        self.assertIsInstance(handled_error, SecurityError)
    
    def test_handle_memory_error(self):
        """测试处理内存错误"""
        original_error = MemoryError("内存不足")
        
        handled_error = self.error_handler.handle_error(original_error)
        
        self.assertIsInstance(handled_error, PerformanceError)
    
    def test_error_statistics(self):
        """测试错误统计"""
        # 处理几个不同类型的错误
        self.error_handler.handle_error(ValueError("错误1"))
        self.error_handler.handle_error(FileNotFoundError("错误2"))
        self.error_handler.handle_error(MemoryError("错误3"))
        
        stats = self.error_handler.get_error_summary()
        
        self.assertEqual(stats['total_errors'], 3)
        self.assertIn('ParsingError', stats['by_type'])
        self.assertIn('PerformanceError', stats['by_type'])
        self.assertTrue(stats['by_severity']['medium'] > 0)
    
    def test_reset_stats(self):
        """测试重置统计"""
        self.error_handler.handle_error(ValueError("错误"))
        self.error_handler.reset_stats()
        
        stats = self.error_handler.get_error_summary()
        self.assertEqual(stats['total_errors'], 0)


class TestErrorDecorator(unittest.TestCase):
    """测试错误处理装饰器"""
    
    def test_error_handler_decorator_success(self):
        """测试装饰器成功执行"""
        @error_handler(operation="测试操作")
        def test_function():
            return "成功"
        
        result = test_function()
        self.assertEqual(result, "成功")
    
    def test_error_handler_decorator_exception(self):
        """测试装饰器异常处理"""
        @error_handler(operation="测试操作", reraise=True)
        def test_function():
            raise ValueError("测试异常")
        
        with self.assertRaises(MCPSheetParserError):
            test_function()
    
    def test_error_handler_decorator_no_reraise(self):
        """测试装饰器不重新抛出异常"""
        @error_handler(operation="测试操作", reraise=False)
        def test_function():
            raise ValueError("测试异常")
        
        result = test_function()
        self.assertIsNone(result)


class TestSafeExecute(unittest.TestCase):
    """测试安全执行函数"""
    
    def test_safe_execute_success(self):
        """测试安全执行成功"""
        def test_func(x, y):
            return x + y
        
        result = safe_execute(test_func, 1, 2)
        self.assertEqual(result, 3)
    
    def test_safe_execute_exception(self):
        """测试安全执行异常"""
        def test_func():
            raise ValueError("测试异常")
        
        result = safe_execute(test_func, default_value="默认值")
        self.assertEqual(result, "默认值")
    
    def test_safe_execute_with_kwargs(self):
        """测试安全执行带关键字参数"""
        def test_func(x, y=10):
            return x * y
        
        result = safe_execute(test_func, 5, y=3)
        self.assertEqual(result, 15)


class TestGlobalErrorHandler(unittest.TestCase):
    """测试全局错误处理器"""
    
    def test_global_error_handler_exists(self):
        """测试全局错误处理器存在"""
        self.assertIsInstance(global_error_handler, ErrorHandler)
    
    def test_global_error_handler_functionality(self):
        """测试全局错误处理器功能"""
        error = ValueError("全局测试错误")
        
        handled_error = global_error_handler.handle_error(error)
        
        self.assertIsInstance(handled_error, ParsingError)
        self.assertEqual(handled_error.original_error, error)


class TestIntegration(unittest.TestCase):
    """集成测试"""
    
    def test_complex_error_scenario(self):
        """测试复杂错误场景"""
        context = create_error_context(
            "复杂测试",
            file_path="/test/complex.xlsx",
            sheet_name="Sheet1",
            cell_position="B5"
        )
        
        # 创建一个有原始错误的自定义异常
        original_error = ValueError("原始异常")
        custom_error = CellProcessingError(
            "B5", 
            "处理失败",
            context=context,
            original_error=original_error
        )
        
        # 验证错误信息
        detailed_msg = custom_error.get_detailed_message()
        self.assertIn("B5", detailed_msg)
        self.assertIn("处理失败", detailed_msg)
        self.assertIn("/test/complex.xlsx", detailed_msg)
        self.assertIn("Sheet1", detailed_msg)
        self.assertIn("ValueError", detailed_msg)
        
        # 验证字典转换
        error_dict = custom_error.to_dict()
        self.assertEqual(error_dict['error_type'], 'CellProcessingError')
        self.assertEqual(error_dict['severity'], 'low')
        self.assertIsNotNone(error_dict['context'])
        self.assertIn("ValueError", error_dict['original_error'])
    
    def test_error_handler_with_context(self):
        """测试带上下文的错误处理器"""
        handler = ErrorHandler()
        context = create_error_context(
            "集成测试",
            file_path="/test/integration.xlsx"
        )
        
        # 处理多种错误类型
        errors = [
            ValueError("值错误"),
            FileNotFoundError("文件未找到"),
            PermissionError("权限错误"),
            MemoryError("内存错误")
        ]
        
        handled_errors = []
        for error in errors:
            handled_error = handler.handle_error(error, context)
            handled_errors.append(handled_error)
        
        # 验证错误类型转换正确
        self.assertIsInstance(handled_errors[0], ParsingError)
        self.assertIsInstance(handled_errors[1], FileNotFoundError)
        self.assertIsInstance(handled_errors[2], SecurityError)
        self.assertIsInstance(handled_errors[3], PerformanceError)
        
        # 验证统计信息
        stats = handler.get_error_summary()
        self.assertEqual(stats['total_errors'], 4)
        self.assertTrue(stats['by_severity']['high'] >= 2)  # FileNotFoundError, SecurityError
        self.assertTrue(stats['by_severity']['medium'] >= 2)  # ParsingError, PerformanceError


if __name__ == '__main__':
    unittest.main() 