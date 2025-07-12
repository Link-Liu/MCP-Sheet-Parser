# test_alignment_enhancement.py
# 测试增强的对齐方式支持

import unittest
from mcp_sheet_parser.config.style import StyleConfig
from mcp_sheet_parser.html_converter import HTMLConverter
from mcp_sheet_parser.style_manager import StyleManager
from mcp_sheet_parser.config import Config


class TestAlignmentEnhancement(unittest.TestCase):
    """测试增强的对齐方式支持"""
    
    def setUp(self):
        """设置测试环境"""
        self.config = Config()
        self.style_config = StyleConfig()
        self.html_converter = HTMLConverter(config=self.config)
        self.style_manager = StyleManager(config=self.config)
    
    def test_chinese_alignment_support(self):
        """测试中文对齐方式支持"""
        # 测试水平对齐
        self.assertEqual(self.style_config.get_alignment_style('左对齐'), 'left')
        self.assertEqual(self.style_config.get_alignment_style('居中'), 'center')
        self.assertEqual(self.style_config.get_alignment_style('右对齐'), 'right')
        self.assertEqual(self.style_config.get_alignment_style('两端对齐'), 'justify')
        self.assertEqual(self.style_config.get_alignment_style('分散对齐'), 'justify')
        self.assertEqual(self.style_config.get_alignment_style('填充'), 'left')
        self.assertEqual(self.style_config.get_alignment_style('常规'), 'left')
        
        # 测试垂直对齐
        self.assertEqual(self.style_config.get_vertical_alignment_style('顶端对齐'), 'top')
        self.assertEqual(self.style_config.get_vertical_alignment_style('垂直居中'), 'middle')
        self.assertEqual(self.style_config.get_vertical_alignment_style('底端对齐'), 'bottom')
        self.assertEqual(self.style_config.get_vertical_alignment_style('垂直两端对齐'), 'middle')
        self.assertEqual(self.style_config.get_vertical_alignment_style('垂直分散对齐'), 'middle')
    
    def test_english_alias_support(self):
        """测试英文别名支持"""
        # 测试水平对齐别名
        self.assertEqual(self.style_config.get_alignment_style('start'), 'left')
        self.assertEqual(self.style_config.get_alignment_style('end'), 'right')
        self.assertEqual(self.style_config.get_alignment_style('middle'), 'center')
        self.assertEqual(self.style_config.get_alignment_style('justified'), 'justify')
        self.assertEqual(self.style_config.get_alignment_style('distribute'), 'justify')
        
        # 测试垂直对齐别名
        self.assertEqual(self.style_config.get_vertical_alignment_style('start'), 'top')
        self.assertEqual(self.style_config.get_vertical_alignment_style('end'), 'bottom')
        self.assertEqual(self.style_config.get_vertical_alignment_style('baseline'), 'top')
    
    def test_numeric_code_support(self):
        """测试数字代码支持"""
        # 测试水平对齐数字代码
        self.assertEqual(self.style_config.get_alignment_style('0'), 'general')
        self.assertEqual(self.style_config.get_alignment_style('1'), 'left')
        self.assertEqual(self.style_config.get_alignment_style('2'), 'center')
        self.assertEqual(self.style_config.get_alignment_style('3'), 'right')
        self.assertEqual(self.style_config.get_alignment_style('4'), 'fill')
        self.assertEqual(self.style_config.get_alignment_style('5'), 'justify')
        self.assertEqual(self.style_config.get_alignment_style('6'), 'centerContinuous')
        self.assertEqual(self.style_config.get_alignment_style('7'), 'distributed')
        
        # 测试垂直对齐数字代码
        self.assertEqual(self.style_config.get_vertical_alignment_style('0'), 'top')
        self.assertEqual(self.style_config.get_vertical_alignment_style('1'), 'center')
        self.assertEqual(self.style_config.get_vertical_alignment_style('2'), 'bottom')
        self.assertEqual(self.style_config.get_vertical_alignment_style('3'), 'justify')
        self.assertEqual(self.style_config.get_vertical_alignment_style('4'), 'distributed')
    
    def test_case_insensitive_support(self):
        """测试大小写不敏感支持"""
        # 测试水平对齐大小写不敏感
        self.assertEqual(self.style_config.get_alignment_style('LEFT'), 'left')
        self.assertEqual(self.style_config.get_alignment_style('Center'), 'center')
        self.assertEqual(self.style_config.get_alignment_style('RIGHT'), 'right')
        self.assertEqual(self.style_config.get_alignment_style('JUSTIFY'), 'justify')
        
        # 测试垂直对齐大小写不敏感
        self.assertEqual(self.style_config.get_vertical_alignment_style('TOP'), 'top')
        self.assertEqual(self.style_config.get_vertical_alignment_style('Center'), 'middle')
        self.assertEqual(self.style_config.get_vertical_alignment_style('BOTTOM'), 'bottom')
    
    def test_edge_cases(self):
        """测试边界情况"""
        # 测试None和空字符串
        self.assertEqual(self.style_config.get_alignment_style(None), 'left')
        self.assertEqual(self.style_config.get_alignment_style(''), 'left')
        self.assertEqual(self.style_config.get_alignment_style('   '), 'left')
        
        self.assertEqual(self.style_config.get_vertical_alignment_style(None), 'top')
        self.assertEqual(self.style_config.get_vertical_alignment_style(''), 'top')
        self.assertEqual(self.style_config.get_vertical_alignment_style('   '), 'top')
        
        # 测试未知对齐方式
        self.assertEqual(self.style_config.get_alignment_style('unknown'), 'left')
        self.assertEqual(self.style_config.get_vertical_alignment_style('unknown'), 'top')
    
    def test_validation_methods(self):
        """测试验证方法"""
        # 测试有效对齐方式
        self.assertTrue(self.style_config.is_valid_alignment('left'))
        self.assertTrue(self.style_config.is_valid_alignment('center'))
        self.assertTrue(self.style_config.is_valid_alignment('右对齐'))
        self.assertTrue(self.style_config.is_valid_alignment('1'))
        
        # 测试无效对齐方式
        self.assertFalse(self.style_config.is_valid_alignment('invalid'))
        self.assertFalse(self.style_config.is_valid_alignment(None))
        self.assertFalse(self.style_config.is_valid_alignment(''))
        
        # 测试有效垂直对齐方式
        self.assertTrue(self.style_config.is_valid_vertical_alignment('top'))
        self.assertTrue(self.style_config.is_valid_vertical_alignment('middle'))
        self.assertTrue(self.style_config.is_valid_vertical_alignment('垂直居中'))
        self.assertTrue(self.style_config.is_valid_vertical_alignment('1'))
        
        # 测试无效垂直对齐方式
        self.assertFalse(self.style_config.is_valid_vertical_alignment('invalid'))
        self.assertFalse(self.style_config.is_valid_vertical_alignment(None))
        self.assertFalse(self.style_config.is_valid_vertical_alignment(''))
    
    def test_supported_alignments(self):
        """测试支持的对齐方式列表"""
        # 测试水平对齐方式列表
        supported_horizontal = self.style_config.get_supported_alignments()
        expected_horizontal = ['left', 'center', 'right', 'justify', 'general', 'distributed', 'centerContinuous', 'fill']
        self.assertEqual(set(supported_horizontal), set(expected_horizontal))
        
        # 测试垂直对齐方式列表
        supported_vertical = self.style_config.get_supported_vertical_alignments()
        expected_vertical = ['top', 'middle', 'bottom']
        self.assertEqual(set(supported_vertical), set(expected_vertical))
    
    def test_html_converter_alignment(self):
        """测试HTML转换器中的对齐方式处理"""
        # 测试中文对齐方式
        style_info = {'align': '居中'}
        styles = self.html_converter._apply_cell_styles(style_info)
        self.assertIn('text-align: center', styles)
        
        # 测试数字代码
        style_info = {'align': '2'}
        styles = self.html_converter._apply_cell_styles(style_info)
        self.assertIn('text-align: center', styles)
        
        # 测试垂直对齐
        style_info = {'valign': '垂直居中'}
        styles = self.html_converter._apply_cell_styles(style_info)
        self.assertIn('vertical-align: middle', styles)
        
        # 测试组合对齐
        style_info = {'align': '右对齐', 'valign': '底端对齐'}
        styles = self.html_converter._apply_cell_styles(style_info)
        self.assertIn('text-align: right', styles)
        self.assertIn('vertical-align: bottom', styles)
    
    def test_style_manager_alignment(self):
        """测试样式管理器中的对齐方式处理"""
        # 测试中文对齐方式
        style = {'align': '居中'}
        css_props = self.style_manager._convert_style_to_css_dict(style)
        self.assertEqual(css_props['text-align'], 'center')
        
        # 测试数字代码
        style = {'align': '2'}
        css_props = self.style_manager._convert_style_to_css_dict(style)
        self.assertEqual(css_props['text-align'], 'center')
        
        # 测试垂直对齐
        style = {'valign': '垂直居中'}
        css_props = self.style_manager._convert_style_to_css_dict(style)
        self.assertEqual(css_props['vertical-align'], 'middle')
        
        # 测试组合对齐
        style = {'align': '右对齐', 'valign': '底端对齐'}
        css_props = self.style_manager._convert_style_to_css_dict(style)
        self.assertEqual(css_props['text-align'], 'right')
        self.assertEqual(css_props['vertical-align'], 'bottom')
    
    def test_excel_specific_alignments(self):
        """测试Excel特定对齐方式"""
        # 测试Excel特定对齐方式
        excel_alignments = [
            'centerContinuous',
            'distributed',
            'fill',
            'general'
        ]
        
        for alignment in excel_alignments:
            result = self.style_config.get_alignment_style(alignment)
            self.assertIsNotNone(result)
            self.assertIn(result, ['left', 'center', 'right', 'justify'])
    
    def test_comprehensive_alignment_mapping(self):
        """测试全面的对齐方式映射"""
        # 测试所有支持的水平对齐方式
        test_cases = [
            ('left', 'left'),
            ('center', 'center'),
            ('right', 'right'),
            ('justify', 'justify'),
            ('左对齐', 'left'),
            ('居中', 'center'),
            ('右对齐', 'right'),
            ('两端对齐', 'justify'),
            ('start', 'left'),
            ('end', 'right'),
            ('middle', 'center'),
            ('0', 'general'),
            ('1', 'left'),
            ('2', 'center'),
            ('3', 'right'),
        ]
        
        for input_alignment, expected_output in test_cases:
            with self.subTest(input_alignment=input_alignment):
                result = self.style_config.get_alignment_style(input_alignment)
                self.assertEqual(result, expected_output)
        
        # 测试所有支持的垂直对齐方式
        test_cases_vertical = [
            ('top', 'top'),
            ('center', 'middle'),
            ('middle', 'middle'),
            ('bottom', 'bottom'),
            ('顶端对齐', 'top'),
            ('垂直居中', 'middle'),
            ('底端对齐', 'bottom'),
            ('start', 'top'),
            ('end', 'bottom'),
            ('0', 'top'),
            ('1', 'center'),
            ('2', 'bottom'),
        ]
        
        for input_alignment, expected_output in test_cases_vertical:
            with self.subTest(input_alignment=input_alignment):
                result = self.style_config.get_vertical_alignment_style(input_alignment)
                self.assertEqual(result, expected_output)


if __name__ == '__main__':
    unittest.main() 