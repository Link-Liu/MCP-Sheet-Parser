# test_parser.py
# parser模块单元测试

import unittest
import tempfile
import os
import pandas as pd
import openpyxl

from mcp_sheet_parser.parser import SheetParser, clean_cell_value
from mcp_sheet_parser.config import Config


class TestSheetParser(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.config = Config()
        
    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir)

    def test_clean_cell_value(self):
        """测试单元格值清理"""
        self.assertEqual(clean_cell_value(None), '')
        self.assertEqual(clean_cell_value(''), '')
        self.assertEqual(clean_cell_value('  test  '), 'test')
        self.assertEqual(clean_cell_value(123), '123')
        self.assertEqual(clean_cell_value(123.45), '123.45')

    def test_unsupported_format(self):
        """测试不支持的文件格式"""
        test_file = os.path.join(self.temp_dir, 'test.txt')
        with open(test_file, 'w') as f:
            f.write('test')
        
        with self.assertRaises(ValueError) as context:
            SheetParser(test_file, self.config)
        
        self.assertIn('不支持的文件格式', str(context.exception))

    def test_csv_parsing(self):
        """测试CSV文件解析"""
        # 创建测试CSV文件
        csv_file = os.path.join(self.temp_dir, 'test.csv')
        csv_data = [
            ['Name', 'Age', 'City'],
            ['John', '25', 'New York'],
            ['Jane', '30', 'London']
        ]
        
        df = pd.DataFrame(csv_data)
        df.to_csv(csv_file, index=False, header=False)
        
        parser = SheetParser(csv_file, self.config)
        sheets = parser.parse()
        
        self.assertEqual(len(sheets), 1)
        sheet = sheets[0]
        
        self.assertEqual(sheet['rows'], 3)
        self.assertEqual(sheet['cols'], 3)
        self.assertEqual(sheet['data'][0][0], 'Name')
        self.assertEqual(sheet['data'][1][0], 'John')
        self.assertEqual(len(sheet['merged_cells']), 0)  # CSV无合并单元格

    def test_excel_parsing(self):
        """测试Excel文件解析"""
        # 创建测试Excel文件
        excel_file = os.path.join(self.temp_dir, 'test.xlsx')
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Test Sheet"
        
        # 添加数据
        data = [
            ['Name', 'Age', 'City'],
            ['John', 25, 'New York'],
            ['Jane', 30, 'London']
        ]
        
        for row_idx, row in enumerate(data, 1):
            for col_idx, value in enumerate(row, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)
        
        wb.save(excel_file)
        
        parser = SheetParser(excel_file, self.config)
        sheets = parser.parse()
        
        self.assertEqual(len(sheets), 1)
        sheet = sheets[0]
        
        self.assertEqual(sheet['sheet_name'], 'Test Sheet')
        self.assertEqual(sheet['rows'], 3)
        self.assertEqual(sheet['cols'], 3)
        self.assertEqual(sheet['data'][0][0], 'Name')

    def test_empty_file(self):
        """测试空文件处理"""
        # 创建空CSV文件
        csv_file = os.path.join(self.temp_dir, 'empty.csv')
        with open(csv_file, 'w') as f:
            pass
        
        parser = SheetParser(csv_file, self.config)
        sheets = parser.parse()
        
        self.assertEqual(len(sheets), 1)
        sheet = sheets[0]
        self.assertEqual(sheet['rows'], 0)
        self.assertEqual(sheet['cols'], 0)
        self.assertEqual(len(sheet['data']), 0)

    def test_file_size_limit(self):
        """测试文件大小限制"""
        # 创建一个超过限制的文件
        large_file = os.path.join(self.temp_dir, 'large.csv')
        
        # 临时设置很小的文件大小限制
        config = Config()
        config.MAX_FILE_SIZE_MB = 0.001  # 1KB
        
        # 写入较大的内容
        with open(large_file, 'w') as f:
            f.write('x' * 2000)  # 2KB
        
        with self.assertRaises(ValueError) as context:
            SheetParser(large_file, config)
        
        self.assertIn('文件过大', str(context.exception))

    def test_nonexistent_file(self):
        """测试不存在的文件"""
        nonexistent_file = os.path.join(self.temp_dir, 'nonexistent.xlsx')
        
        with self.assertRaises(FileNotFoundError):
            SheetParser(nonexistent_file, self.config)


if __name__ == '__main__':
    unittest.main()
