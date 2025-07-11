# test_parser.py
# 测试表格解析模块

import unittest
import tempfile
import os
from pathlib import Path

from mcp_sheet_parser.parser import SheetParser
from mcp_sheet_parser.utils import clean_cell_value
from mcp_sheet_parser.config import Config
from mcp_sheet_parser.utils import setup_logger


class TestSheetParser(unittest.TestCase):
    """测试SheetParser类"""
    
    def setUp(self):
        """测试前准备"""
        self.config = Config()
        self.logger = setup_logger(__name__)
    
    def test_clean_cell_value(self):
        """测试单元格值清理函数"""
        # 测试各种情况
        self.assertEqual(clean_cell_value(None), '')
        self.assertEqual(clean_cell_value(''), '')
        self.assertEqual(clean_cell_value('  test  '), 'test')
        self.assertEqual(clean_cell_value(123), '123')
        self.assertEqual(clean_cell_value(123.45), '123.45')
    
    def test_nonexistent_file(self):
        """测试不存在的文件"""
        with self.assertRaises(FileNotFoundError):
            parser = SheetParser('nonexistent.xlsx', self.config)
    
    def test_unsupported_format(self):
        """测试不支持的文件格式"""
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            tmp.write(b'test content')
            tmp_path = tmp.name
        
        try:
            with self.assertRaises(ValueError):
                parser = SheetParser(tmp_path, self.config)
        finally:
            os.unlink(tmp_path)
    
    def test_csv_parsing(self):
        """测试CSV文件解析"""
        # 创建测试CSV文件
        csv_content = "Name,Age,City\nJohn,25,NYC\nJane,30,LA"
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as tmp:
            tmp.write(csv_content)
            tmp_path = tmp.name
        
        try:
            parser = SheetParser(tmp_path, self.config)
            result = parser.parse()
            
            self.assertIsInstance(result, list)
            self.assertEqual(len(result), 1)  # 一个工作表
            
            sheet = result[0]
            self.assertEqual(sheet['rows'], 3)
            self.assertEqual(sheet['cols'], 3)
            self.assertEqual(len(sheet['data']), 3)
            
        finally:
            os.unlink(tmp_path)
    
    def test_empty_file(self):
        """测试空文件"""
        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as tmp:
            tmp_path = tmp.name
        
        try:
            parser = SheetParser(tmp_path, self.config)
            result = parser.parse()
            
            self.assertIsInstance(result, list)
            self.assertEqual(len(result), 1)
            
            sheet = result[0]
            self.assertEqual(sheet['rows'], 0)
            self.assertEqual(sheet['cols'], 0)
            self.assertEqual(len(sheet['data']), 0)
            
        finally:
            os.unlink(tmp_path)
    
    def test_excel_parsing(self):
        """测试Excel文件解析"""
        # 这个测试需要一个真实的Excel文件
        # 在实际环境中，可以使用项目中的示例文件
        examples_dir = Path(__file__).parent.parent / 'examples'
        excel_files = list(examples_dir.glob('*.xlsx'))
        
        if excel_files:
            test_file = excel_files[0]
            parser = SheetParser(str(test_file), self.config)
            result = parser.parse()
            
            self.assertIsInstance(result, list)
            self.assertGreater(len(result), 0)
            
            for sheet in result:
                self.assertIn('sheet_name', sheet)
                self.assertIn('rows', sheet)
                self.assertIn('cols', sheet)
                self.assertIn('data', sheet)
                self.assertIn('styles', sheet)
    
    def test_file_size_limit(self):
        """测试文件大小限制"""
        # 修改配置以设置较小的文件大小限制
        small_config = Config()
        small_config.MAX_FILE_SIZE_MB = 0.001  # 1KB limit
        
        # 创建一个略大的文件
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tmp:
            tmp.write("A,B,C\n" * 1000)  # 应该超过1KB
            tmp_path = tmp.name
        
        try:
            with self.assertRaises(ValueError):
                parser = SheetParser(tmp_path, small_config)
        finally:
            os.unlink(tmp_path)


if __name__ == '__main__':
    unittest.main()
