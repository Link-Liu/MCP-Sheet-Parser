# test_utils.py
# utils模块单元测试

import unittest
import tempfile
import os
import logging
from pathlib import Path

from mcp_sheet_parser.utils import (
    setup_logger,
    validate_file_path,
    get_file_extension,
    is_supported_format,
    sanitize_filename,
    ensure_output_dir,
    format_file_size,
    batch_process_files,
    get_file_info
)


class TestUtils(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir)

    def test_setup_logger(self):
        """测试日志设置"""
        logger = setup_logger('test_logger', logging.DEBUG)
        self.assertEqual(logger.name, 'test_logger')
        self.assertEqual(logger.level, logging.DEBUG)
        self.assertTrue(len(logger.handlers) > 0)

    def test_validate_file_path(self):
        """测试文件路径验证"""
        # 创建测试文件
        test_file = os.path.join(self.temp_dir, 'test.txt')
        with open(test_file, 'w') as f:
            f.write('test')
        
        # 有效文件
        self.assertTrue(validate_file_path(test_file))
        
        # 无效文件
        self.assertFalse(validate_file_path('/nonexistent/file.txt'))
        
        # 目录不是文件
        self.assertFalse(validate_file_path(self.temp_dir))

    def test_get_file_extension(self):
        """测试文件扩展名获取"""
        self.assertEqual(get_file_extension('test.xlsx'), '.xlsx')
        self.assertEqual(get_file_extension('test.CSV'), '.csv')
        self.assertEqual(get_file_extension('path/to/file.XLS'), '.xls')
        self.assertEqual(get_file_extension('noextension'), '')

    def test_is_supported_format(self):
        """测试支持格式检查"""
        self.assertTrue(is_supported_format('test.xlsx'))
        self.assertTrue(is_supported_format('test.xls'))
        self.assertTrue(is_supported_format('test.csv'))
        self.assertTrue(is_supported_format('test.et'))
        self.assertTrue(is_supported_format('test.xltx'))  # 新增的格式
        self.assertFalse(is_supported_format('test.txt'))
        self.assertFalse(is_supported_format('test.doc'))

    def test_sanitize_filename(self):
        """测试文件名清理"""
        test_cases = [
            ('normal_file.txt', 'normal_file.txt'),
            ('file<>with|special?chars*.txt', 'file__with_special_chars_.txt'),
            ('file\x00with\x1fcontrol.txt', 'filewithcontrol.txt')
        ]
        
        for input_name, expected in test_cases:
            result = sanitize_filename(input_name)
            self.assertEqual(result, expected)
        
        # 测试长文件名
        long_name = 'very_long_filename_' + 'x' * 300 + '.txt'
        result = sanitize_filename(long_name)
        self.assertLessEqual(len(result), 255)

    def test_ensure_output_dir(self):
        """测试输出目录创建"""
        output_path = os.path.join(self.temp_dir, 'subdir', 'file.html')
        
        # 目录不存在
        self.assertFalse(os.path.exists(os.path.dirname(output_path)))
        
        # 创建目录
        ensure_output_dir(output_path)
        
        # 目录应该存在
        self.assertTrue(os.path.exists(os.path.dirname(output_path)))

    def test_format_file_size(self):
        """测试文件大小格式化"""
        test_cases = [
            (0, '0B'),
            (512, '512B'),
            (1024, '1.0KB'),
            (1536, '1.5KB'),
            (1048576, '1.0MB'),
            (1073741824, '1.0GB'),
        ]
        
        for size_bytes, expected in test_cases:
            result = format_file_size(size_bytes)
            self.assertEqual(result, expected)

    def test_get_file_info(self):
        """测试文件信息获取"""
        # 创建测试文件
        test_file = os.path.join(self.temp_dir, 'test.xlsx')
        test_content = 'test content'
        with open(test_file, 'w') as f:
            f.write(test_content)
        
        info = get_file_info(test_file)
        
        self.assertIn('path', info)
        self.assertIn('name', info)
        self.assertIn('size', info)
        self.assertIn('size_formatted', info)
        self.assertIn('extension', info)
        self.assertIn('is_supported', info)
        self.assertIn('modified_time', info)
        
        self.assertEqual(info['name'], 'test.xlsx')
        self.assertEqual(info['extension'], '.xlsx')
        self.assertTrue(info['is_supported'])

    def test_get_file_info_nonexistent(self):
        """测试不存在文件的信息获取"""
        info = get_file_info('/nonexistent/file.txt')
        
        self.assertIn('error', info)
        self.assertEqual(info['error'], '文件不存在')

    def test_batch_process_files(self):
        """测试批量文件处理"""
        # 创建测试文件
        test_files = []
        for i in range(3):
            test_file = os.path.join(self.temp_dir, f'test{i}.txt')
            with open(test_file, 'w') as f:
                f.write(f'content {i}')
            test_files.append(test_file)
        
        def simple_processor(file_path):
            with open(file_path, 'r') as f:
                return f.read()
        
        results = batch_process_files(test_files, simple_processor)
        
        self.assertEqual(len(results), 3)
        for i, result in enumerate(results):
            self.assertTrue(result['success'])
            self.assertEqual(result['result'], f'content {i}')

    def test_batch_process_with_errors(self):
        """测试带错误的批量处理"""
        test_files = ['nonexistent1.txt', 'nonexistent2.txt']
        
        def failing_processor(file_path):
            raise ValueError(f"Cannot process {file_path}")
        
        results = batch_process_files(test_files, failing_processor)
        
        self.assertEqual(len(results), 2)
        for result in results:
            self.assertFalse(result['success'])
            self.assertIn('error', result)


if __name__ == '__main__':
    unittest.main()
