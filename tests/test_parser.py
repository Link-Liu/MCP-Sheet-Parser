# test_parser.py
# SheetParser 单元测试

import unittest
import os
import tempfile
import pandas as pd
import openpyxl
from mcp_sheet_parser.parser import SheetParser
from pandas.errors import EmptyDataError

class TestSheetParser(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir)

    def test_csv_parsing(self):
        """测试CSV文件解析功能"""
        # 创建测试CSV文件
        csv_data = [
            ['姓名', '年龄', '城市'],
            ['张三', '25', '北京'],
            ['李四', '30', '上海'],
            ['王五', '28', '广州']
        ]
        csv_path = os.path.join(self.temp_dir, 'test.csv')
        df = pd.DataFrame(csv_data)
        df.to_csv(csv_path, index=False, header=False)
        
        # 解析CSV文件
        parser = SheetParser(csv_path)
        result = parser.parse()
        
        # 验证结果
        self.assertEqual(len(result), 1)
        sheet = result[0]
        self.assertEqual(sheet['sheet_name'], 'test.csv')
        self.assertEqual(sheet['rows'], 4)
        self.assertEqual(sheet['cols'], 3)
        self.assertEqual(sheet['data'], csv_data)
        self.assertEqual(sheet['merged_cells'], [])

    def test_excel_parsing(self):
        """测试Excel文件解析功能"""
        # 创建测试Excel文件
        excel_path = os.path.join(self.temp_dir, 'test.xlsx')
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Sheet1"
        ws1['A1'] = '产品'
        ws1['B1'] = '价格'
        ws1['C1'] = '库存'
        ws1['A2'] = '苹果'
        ws1['B2'] = 5.5
        ws1['C2'] = 100
        ws1['A3'] = '香蕉'
        ws1['B3'] = 3.2
        ws1['C3'] = 200
        
        # 添加合并单元格
        ws1.merge_cells('A1:C1')
        
        # 创建第二个工作表
        ws2 = wb.create_sheet("Sheet2")
        ws2['A1'] = '部门'
        ws2['B1'] = '人数'
        ws2['A2'] = '技术部'
        ws2['B2'] = 50
        
        wb.save(excel_path)
        
        # 解析Excel文件
        parser = SheetParser(excel_path)
        result = parser.parse()
        
        # 验证结果
        self.assertEqual(len(result), 2)
        
        # 验证第一个工作表
        sheet1 = result[0]
        self.assertEqual(sheet1['sheet_name'], 'Sheet1')
        self.assertEqual(sheet1['rows'], 3)
        self.assertEqual(sheet1['cols'], 3)
        self.assertEqual(len(sheet1['merged_cells']), 1)
        self.assertEqual(sheet1['merged_cells'][0], (0, 0, 0, 2))  # A1:C1合并
        
        # 验证第二个工作表
        sheet2 = result[1]
        self.assertEqual(sheet2['sheet_name'], 'Sheet2')
        self.assertEqual(sheet2['rows'], 2)
        self.assertEqual(sheet2['cols'], 2)

    def test_unsupported_format(self):
        """测试不支持的文件格式"""
        txt_path = os.path.join(self.temp_dir, 'test.txt')
        with open(txt_path, 'w') as f:
            f.write('test content')
        
        parser = SheetParser(txt_path)
        with self.assertRaises(ValueError):
            parser.parse()

    def test_empty_file(self):
        """测试空文件处理"""
        # 创建空CSV文件
        csv_path = os.path.join(self.temp_dir, 'empty.csv')
        with open(csv_path, 'w') as f:
            pass
        
        parser = SheetParser(csv_path)
        result = parser.parse()
        
        self.assertEqual(len(result), 1)
        sheet = result[0]
        self.assertEqual(sheet['rows'], 0)
        self.assertEqual(sheet['cols'], 0)
        self.assertEqual(sheet['data'], [])

if __name__ == '__main__':
    unittest.main()
