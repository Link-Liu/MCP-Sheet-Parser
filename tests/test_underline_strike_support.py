# test_underline_strike_support.py
# 下划线和删除线支持测试

import unittest
import tempfile
import os
from unittest.mock import Mock, patch

from mcp_sheet_parser.style_manager import StyleManager
from mcp_sheet_parser.config import Config


class TestUnderlineStrikeSupport(unittest.TestCase):
    """下划线和删除线支持测试"""
    
    def setUp(self):
        """设置测试环境"""
        self.config = Config()
        self.style_manager = StyleManager(self.config)
        
        # 测试数据
        self.test_sheet_data = {
            'data': [
                ['标题', '内容', '备注'],
                ['正常文本', '下划线文本', '删除线文本'],
                ['粗体文本', '斜体文本', '组合样式']
            ],
            'styles': [
                [
                    {'bold': True, 'bg_color': '#f0f0f0'},
                    {'bold': True, 'bg_color': '#f0f0f0'},
                    {'bold': True, 'bg_color': '#f0f0f0'}
                ],
                [
                    {},
                    {'underline': True},
                    {'strike': True}
                ],
                [
                    {'bold': True},
                    {'italic': True},
                    {'bold': True, 'underline': True, 'strike': True}
                ]
            ]
        }

    def test_underline_style_conversion(self):
        """测试下划线样式转换"""
        # 测试单独的下划线
        style = {'underline': True}
        css_props = self.style_manager._convert_style_to_css_dict(style)
        
        self.assertIn('text-decoration', css_props)
        self.assertEqual(css_props['text-decoration'], 'underline')

    def test_strike_style_conversion(self):
        """测试删除线样式转换"""
        # 测试单独的删除线
        style = {'strike': True}
        css_props = self.style_manager._convert_style_to_css_dict(style)
        
        self.assertIn('text-decoration', css_props)
        self.assertEqual(css_props['text-decoration'], 'line-through')

    def test_combined_underline_strike(self):
        """测试下划线和删除线组合"""
        # 测试同时有下划线和删除线
        style = {'underline': True, 'strike': True}
        css_props = self.style_manager._convert_style_to_css_dict(style)
        
        self.assertIn('text-decoration', css_props)
        self.assertEqual(css_props['text-decoration'], 'underline line-through')

    def test_underline_with_other_styles(self):
        """测试下划线与其他样式组合"""
        style = {
            'bold': True,
            'italic': True,
            'underline': True,
            'font_color': '#ff0000'
        }
        css_props = self.style_manager._convert_style_to_css_dict(style)
        
        self.assertIn('font-weight', css_props)
        self.assertIn('font-style', css_props)
        self.assertIn('text-decoration', css_props)
        self.assertIn('color', css_props)
        
        self.assertEqual(css_props['font-weight'], 'bold')
        self.assertEqual(css_props['font-style'], 'italic')
        self.assertEqual(css_props['text-decoration'], 'underline')
        self.assertEqual(css_props['color'], '#ff0000')

    def test_strike_with_other_styles(self):
        """测试删除线与其他样式组合"""
        style = {
            'bold': True,
            'strike': True,
            'bg_color': '#ffff00'
        }
        css_props = self.style_manager._convert_style_to_css_dict(style)
        
        self.assertIn('font-weight', css_props)
        self.assertIn('text-decoration', css_props)
        self.assertIn('background-color', css_props)
        
        self.assertEqual(css_props['font-weight'], 'bold')
        self.assertEqual(css_props['text-decoration'], 'line-through')
        self.assertEqual(css_props['background-color'], '#ffff00')

    def test_style_description_with_underline_strike(self):
        """测试样式描述包含下划线和删除线"""
        # 测试下划线描述
        style = {'underline': True}
        description = self.style_manager._generate_style_description(style)
        self.assertIn('下划线', description)
        
        # 测试删除线描述
        style = {'strike': True}
        description = self.style_manager._generate_style_description(style)
        self.assertIn('删除线', description)
        
        # 测试组合描述
        style = {'underline': True, 'strike': True, 'bold': True}
        description = self.style_manager._generate_style_description(style)
        self.assertIn('下划线', description)
        self.assertIn('删除线', description)
        self.assertIn('粗体', description)

    def test_semantic_class_names_with_underline_strike(self):
        """测试语义化类名包含下划线和删除线"""
        # 测试下划线类名
        style = {'underline': True}
        class_name = self.style_manager._generate_semantic_class_name(
            style, [['测试']], [(0, 0)]
        )
        self.assertIn('underline', class_name)
        
        # 测试删除线类名
        style = {'strike': True}
        class_name = self.style_manager._generate_semantic_class_name(
            style, [['测试']], [(0, 0)]
        )
        self.assertIn('strike', class_name)
        
        # 测试组合类名
        style = {'underline': True, 'strike': True, 'bold': True}
        class_name = self.style_manager._generate_semantic_class_name(
            style, [['测试']], [(0, 0)]
        )
        self.assertIn('underline', class_name)
        self.assertIn('strike', class_name)
        self.assertIn('bold', class_name)

    def test_css_generation_with_underline_strike(self):
        """测试CSS生成包含下划线和删除线"""
        css_content, cell_class_map = self.style_manager.generate_css_classes(
            self.test_sheet_data, 
            use_semantic_names=True,
            min_usage_threshold=1
        )
        
        # 检查CSS内容包含text-decoration
        self.assertIn('text-decoration', css_content)
        
        # 检查生成了相应的CSS类
        self.assertGreater(len(self.style_manager.style_classes), 0)
        
        # 检查是否有包含下划线或删除线的类
        found_underline_strike = False
        for style_class in self.style_manager.style_classes.values():
            if 'text-decoration' in style_class.css_properties:
                found_underline_strike = True
                break
        
        self.assertTrue(found_underline_strike)

    def test_style_config_underline_strike_support(self):
        """测试样式配置支持下划线和删除线"""
        from mcp_sheet_parser.config.style import StyleConfig
        
        style_config = StyleConfig()
        
        # 测试字体样式创建
        font_style = style_config.create_font_style(
            underline=True, strike=False
        )
        self.assertIn('text-decoration', font_style)
        self.assertEqual(font_style['text-decoration'], 'underline')
        
        # 测试删除线
        font_style = style_config.create_font_style(
            underline=False, strike=True
        )
        self.assertIn('text-decoration', font_style)
        self.assertEqual(font_style['text-decoration'], 'line-through')
        
        # 测试组合
        font_style = style_config.create_font_style(
            underline=True, strike=True
        )
        self.assertIn('text-decoration', font_style)
        self.assertEqual(font_style['text-decoration'], 'underline line-through')
        
        # 测试完整单元格样式
        cell_style = style_config.create_cell_style(
            underline=True, strike=True, font_color='red'
        )
        self.assertIn('text-decoration', cell_style)
        self.assertEqual(cell_style['text-decoration'], 'underline line-through')

    def test_no_underline_strike_when_not_set(self):
        """测试未设置下划线删除线时不生成相关CSS"""
        style = {'bold': True, 'font_color': '#000000'}
        css_props = self.style_manager._convert_style_to_css_dict(style)
        
        # 不应该包含text-decoration
        self.assertNotIn('text-decoration', css_props)
        
        # 应该包含其他样式
        self.assertIn('font-weight', css_props)
        self.assertIn('color', css_props)

    def test_underline_strike_in_conditional_formatting(self):
        """测试条件格式化中的下划线和删除线"""
        from mcp_sheet_parser.style_manager import ConditionalRule, ConditionalType, ComparisonOperator
        
        # 创建包含下划线的条件规则
        underline_rule = ConditionalRule(
            name="下划线标记",
            type=ConditionalType.VALUE_RANGE,
            operator=ComparisonOperator.GREATER_THAN,
            values=[0],
            styles={'text-decoration': 'underline', 'color': '#ff0000'},
            priority=10
        )
        
        self.style_manager.add_conditional_rule(underline_rule)
        
        # 创建测试数据
        test_data = {
            'data': [['10'], ['-5'], ['20']],
            'styles': [[{}], [{}], [{}]]
        }
        
        # 应用条件格式化
        conditional_styles = self.style_manager.apply_conditional_formatting(test_data)
        
        # 检查是否应用了下划线样式
        found_underline = False
        for style in conditional_styles.values():
            if 'text-decoration' in style and 'underline' in style:
                found_underline = True
                break
        
        self.assertTrue(found_underline)


if __name__ == '__main__':
    unittest.main() 