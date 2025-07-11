# test_html_converter.py
# HTML转换器测试

import unittest
import tempfile
import os

from mcp_sheet_parser.html_converter import HTMLConverter
from mcp_sheet_parser.config import Config, THEMES


class TestHTMLConverter(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.sample_sheet_data = {
            'sheet_name': 'Test Sheet',
            'rows': 3,
            'cols': 3,
            'data': [
                ['Name', 'Age', 'City'],
                ['John', '25', 'New York'],
                ['Jane', '30', 'London']
            ],
            'styles': [
                [{}, {}, {}],
                [{}, {}, {}],
                [{}, {}, {}]
            ],
            'merged_cells': [],
            'comments': {},
            'hyperlinks': {}
        }
    
    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir)

    def test_simple_table(self):
        """测试简单表格转换"""
        converter = HTMLConverter(self.sample_sheet_data)
        html = converter.to_html()
        
        self.assertIn('<table', html)
        self.assertIn('Name', html)
        self.assertIn('John', html)
        self.assertIn('Jane', html)
        self.assertIn('</table>', html)

    def test_table_only_output(self):
        """测试只输出表格"""
        converter = HTMLConverter(self.sample_sheet_data)
        html = converter.to_html(table_only=True)
        
        self.assertIn('<table', html)
        self.assertNotIn('<!DOCTYPE html>', html)
        self.assertNotIn('<html>', html)

    def test_different_themes(self):
        """测试不同主题"""
        for theme_name in THEMES.keys():
            converter = HTMLConverter(self.sample_sheet_data, theme=theme_name)
            html = converter.to_html()
            
            self.assertIn('<table', html)
            self.assertIn('Name', html)

    def test_with_merged_cells(self):
        """测试合并单元格"""
        data_with_merged = self.sample_sheet_data.copy()
        data_with_merged['merged_cells'] = [(0, 0, 0, 1)]  # 合并A1:B1
        
        converter = HTMLConverter(data_with_merged)
        html = converter.to_html()
        
        self.assertIn('colspan="2"', html)

    def test_with_comments(self):
        """测试包含注释"""
        data_with_comments = self.sample_sheet_data.copy()
        data_with_comments['comments'] = {'0_0': 'This is a header'}
        
        converter = HTMLConverter(data_with_comments)
        html = converter.to_html()
        
        self.assertIn('title=', html)
        self.assertIn('This is a header', html)

    def test_with_hyperlinks(self):
        """测试包含超链接"""
        data_with_links = self.sample_sheet_data.copy()
        data_with_links['hyperlinks'] = {'1_2': 'https://example.com'}
        
        converter = HTMLConverter(data_with_links)
        html = converter.to_html()
        
        self.assertIn('<a href="https://example.com"', html)

    def test_config_options(self):
        """测试配置选项"""
        config = Config()
        config.INCLUDE_COMMENTS = False
        
        data_with_comments = self.sample_sheet_data.copy()
        data_with_comments['comments'] = {'0_0': 'This is a comment'}
        
        converter = HTMLConverter(data_with_comments, config=config)
        html = converter.to_html()
        
        self.assertNotIn('title=', html)

    def test_export_to_file(self):
        """测试导出到文件"""
        output_path = os.path.join(self.temp_dir, 'test.html')
        
        converter = HTMLConverter(self.sample_sheet_data)
        success = converter.export_to_file(output_path)
        
        self.assertTrue(success)
        self.assertTrue(os.path.exists(output_path))
        
        # 验证文件内容
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('<table', content)
        self.assertIn('Name', content)

    def test_empty_table(self):
        """测试空表格"""
        empty_data = {
            'sheet_name': 'Empty',
            'rows': 0,
            'cols': 0,
            'data': [],
            'styles': [],
            'merged_cells': [],
            'comments': {},
            'hyperlinks': {}
        }
        
        converter = HTMLConverter(empty_data)
        html = converter.to_html()
        
        self.assertIn('表格为空', html)

    def test_html_escaping(self):
        """测试HTML转义"""
        data_with_html = self.sample_sheet_data.copy()
        data_with_html['data'][1][0] = '<script>alert("xss")</script>'
        
        converter = HTMLConverter(data_with_html)
        html = converter.to_html()
        
        # 检查单元格内容是否正确转义
        self.assertIn('&lt;script&gt;alert(&quot;xss&quot;)&lt;', html)
        # 确保没有在单元格中直接包含<script>标签
        self.assertNotIn('<td><script>', html)

    def test_cell_type_detection(self):
        """测试单元格类型检测"""
        converter = HTMLConverter(self.sample_sheet_data)
        
        # 基本功能测试
        html = converter.to_html()
        self.assertIn('<th>', html)  # 表头
        self.assertIn('<td>', html)  # 数据单元格


if __name__ == '__main__':
    unittest.main()
