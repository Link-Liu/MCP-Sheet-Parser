# test_html_converter.py
# HTMLConverter 单元测试

import unittest
from mcp_sheet_parser.html_converter import HTMLConverter

class TestHTMLConverter(unittest.TestCase):
    def test_simple_table(self):
        table_data = {
            'sheet_name': '测试',
            'rows': 2,
            'cols': 3,
            'data': [
                ['姓名', '年龄', '城市'],
                ['张三', '25', '北京']
            ],
            'merged_cells': []
        }
        converter = HTMLConverter(table_data)
        html = converter.to_html()
        self.assertIn('<table', html)
        self.assertIn('<th>姓名</th>', html)
        self.assertIn('<td>张三</td>', html)
        self.assertIn('<td>25</td>', html)
        self.assertIn('<td>北京</td>', html)

    def test_merged_cells(self):
        table_data = {
            'sheet_name': '合并测试',
            'rows': 2,
            'cols': 3,
            'data': [
                ['标题', '', ''],
                ['A', 'B', 'C']
            ],
            'merged_cells': [(0, 0, 0, 2)]  # 第一行A1:C1合并
        }
        converter = HTMLConverter(table_data)
        html = converter.to_html()
        # 检查合并单元格属性（rowspan=1可省略）
        self.assertIn('<th colspan="3">标题</th>', html)
        # 检查被合并的单元格未重复输出
        self.assertNotIn('<th></th>', html)
        self.assertIn('<td>A</td>', html)
        self.assertIn('<td>B</td>', html)
        self.assertIn('<td>C</td>', html)

    def test_empty_table(self):
        table_data = {
            'sheet_name': '空表',
            'rows': 0,
            'cols': 0,
            'data': [],
            'merged_cells': []
        }
        converter = HTMLConverter(table_data)
        html = converter.to_html()
        self.assertIn('<table', html)
        self.assertIn('</table>', html)

if __name__ == '__main__':
    unittest.main()
