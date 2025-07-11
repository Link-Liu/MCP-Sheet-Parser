# test_style_manager.py
# 样式管理器测试

import unittest
import tempfile
import os
from unittest.mock import Mock, patch

from mcp_sheet_parser.style_manager import (
    StyleManager, ConditionalRule, ConditionalType, ComparisonOperator,
    StyleClass, StyleTemplate, PREDEFINED_RULES
)
from mcp_sheet_parser.config import Config


class TestStyleManager(unittest.TestCase):
    """样式管理器测试"""
    
    def setUp(self):
        """设置测试环境"""
        self.config = Config()
        self.style_manager = StyleManager(self.config)
        
        # 测试数据
        self.sample_sheet_data = {
            'data': [
                ['姓名', '年龄', '工资', '状态'],
                ['张三', '25', '5000', '在职'],
                ['李四', '30', '8000', '在职'],
                ['王五', '28', '-2000', '离职'],
                ['赵六', '35', '12000', '在职']
            ],
            'styles': [
                [
                    {'bold': True, 'bg_color': '#f0f0f0'},
                    {'bold': True, 'bg_color': '#f0f0f0'},
                    {'bold': True, 'bg_color': '#f0f0f0'},
                    {'bold': True, 'bg_color': '#f0f0f0'}
                ],
                [
                    {},
                    {'font_color': '#333'},
                    {'font_color': '#333'},
                    {'font_color': '#28a745'}
                ],
                [
                    {},
                    {'font_color': '#333'},
                    {'font_color': '#333'},
                    {'font_color': '#28a745'}
                ],
                [
                    {},
                    {'font_color': '#333'},
                    {'font_color': '#dc3545'},
                    {'font_color': '#dc3545'}
                ],
                [
                    {},
                    {'font_color': '#333'},
                    {'font_color': '#333'},
                    {'font_color': '#28a745'}
                ]
            ]
        }

    def test_css_class_generation(self):
        """测试CSS类生成"""
        css_content, cell_class_map = self.style_manager.generate_css_classes(
            self.sample_sheet_data, 
            use_semantic_names=True,
            min_usage_threshold=2
        )
        
        # 检查CSS内容不为空
        self.assertNotEqual(css_content, "")
        
        # 检查是否生成了CSS类
        self.assertGreater(len(self.style_manager.style_classes), 0)
        
        # 检查单元格类映射
        self.assertIsInstance(cell_class_map, dict)
        
        # 检查统计信息
        stats = self.style_manager.get_style_statistics()
        self.assertIn('total_styles', stats)
        self.assertIn('unique_styles', stats)
        self.assertIn('class_reuse_rate', stats)

    def test_semantic_class_names(self):
        """测试语义化类名生成"""
        # 创建带有特定样式的数据
        test_data = {
            'data': [['标题'], ['内容']],
            'styles': [
                [{'bold': True, 'font_color': '#000000', 'bg_color': '#ffff00'}],
                [{'italic': True, 'font_color': '#ff0000'}]
            ]
        }
        
        css_content, cell_class_map = self.style_manager.generate_css_classes(
            test_data, use_semantic_names=True, min_usage_threshold=1
        )
        
        # 检查是否生成了有意义的类名
        class_names = list(self.style_manager.style_classes.keys())
        self.assertGreater(len(class_names), 0)
        
        # 检查类名包含样式特征
        found_meaningful_name = any(
            any(feature in name for feature in ['bold', 'highlight', 'red', 'header'])
            for name in class_names
        )
        self.assertTrue(found_meaningful_name)

    def test_conditional_formatting_value_range(self):
        """测试数值范围条件格式化"""
        # 添加正值条件规则
        positive_rule = ConditionalRule(
            name="正值标记",
            type=ConditionalType.VALUE_RANGE,
            operator=ComparisonOperator.GREATER_THAN,
            values=[0],
            styles={'color': '#28a745', 'font-weight': 'bold'},
            priority=10
        )
        
        self.style_manager.add_conditional_rule(positive_rule)
        
        # 应用条件格式化
        conditional_styles = self.style_manager.apply_conditional_formatting(
            self.sample_sheet_data
        )
        
        # 检查是否有条件样式应用
        self.assertGreater(len(conditional_styles), 0)
        
        # 检查统计信息
        stats = self.style_manager.get_style_statistics()
        self.assertGreater(stats.get('conditional_rules_applied', 0), 0)

    def test_conditional_formatting_color_scale(self):
        """测试颜色渐变条件格式化"""
        color_rule = ConditionalRule(
            name="颜色渐变",
            type=ConditionalType.COLOR_SCALE,
            operator=ComparisonOperator.BETWEEN,
            values=['#ff0000', '#00ff00'],
            styles={},
            priority=5
        )
        
        self.style_manager.add_conditional_rule(color_rule)
        
        # 创建数值数据
        numeric_data = {
            'data': [['10'], ['20'], ['30'], ['40'], ['50']],
            'styles': [[{}], [{}], [{}], [{}], [{}]]
        }
        
        conditional_styles = self.style_manager.apply_conditional_formatting(numeric_data)
        
        # 检查是否应用了颜色渐变
        self.assertGreater(len(conditional_styles), 0)
        
        # 检查样式包含背景色
        for style in conditional_styles.values():
            self.assertIn('background-color', style)

    def test_conditional_formatting_data_bars(self):
        """测试数据条条件格式化"""
        data_bar_rule = ConditionalRule(
            name="数据条",
            type=ConditionalType.DATA_BARS,
            operator=ComparisonOperator.GREATER_THAN,
            values=['#4472C4'],
            styles={},
            priority=5
        )
        
        self.style_manager.add_conditional_rule(data_bar_rule)
        
        # 创建数值数据
        numeric_data = {
            'data': [['100'], ['200'], ['300']],
            'styles': [[{}], [{}], [{}]]
        }
        
        conditional_styles = self.style_manager.apply_conditional_formatting(numeric_data)
        
        # 检查是否应用了数据条
        self.assertGreater(len(conditional_styles), 0)
        
        # 检查样式包含linear-gradient
        found_gradient = any('linear-gradient' in style for style in conditional_styles.values())
        self.assertTrue(found_gradient)

    def test_conditional_formatting_duplicate_values(self):
        """测试重复值条件格式化"""
        duplicate_rule = ConditionalRule(
            name="重复值",
            type=ConditionalType.DUPLICATE_VALUES,
            operator=ComparisonOperator.EQUAL,
            values=[],
            styles={'background-color': '#fff3cd', 'color': '#856404'},
            priority=15
        )
        
        self.style_manager.add_conditional_rule(duplicate_rule)
        
        # 创建包含重复值的数据
        duplicate_data = {
            'data': [['A'], ['B'], ['A'], ['C'], ['B']],
            'styles': [[{}], [{}], [{}], [{}], [{}]]
        }
        
        conditional_styles = self.style_manager.apply_conditional_formatting(duplicate_data)
        
        # 检查是否标记了重复值
        self.assertGreater(len(conditional_styles), 0)
        
        # 应该标记4个单元格（A出现2次，B出现2次）
        self.assertEqual(len(conditional_styles), 4)

    def test_style_templates(self):
        """测试样式模板"""
        # 测试设置预定义模板
        self.assertTrue(self.style_manager.set_template('business'))
        self.assertIsNotNone(self.style_manager.current_template)
        self.assertEqual(self.style_manager.current_template.name, 'business')
        
        # 测试设置不存在的模板
        self.assertFalse(self.style_manager.set_template('nonexistent'))
        
        # 测试创建自定义模板
        custom_template = StyleTemplate(
            name="custom",
            description="自定义模板",
            global_css=".custom { color: red; }"
        )
        
        self.style_manager.create_custom_template(custom_template)
        self.assertIn('custom', self.style_manager.templates)
        
        # 测试使用自定义模板
        self.assertTrue(self.style_manager.set_template('custom'))

    def test_predefined_rules(self):
        """测试预定义规则"""
        # 检查预定义规则存在
        self.assertIn('financial_positive', PREDEFINED_RULES)
        self.assertIn('financial_negative', PREDEFINED_RULES)
        
        # 测试预定义规则类型
        rule = PREDEFINED_RULES['financial_positive']
        self.assertIsInstance(rule, ConditionalRule)
        self.assertEqual(rule.type, ConditionalType.VALUE_RANGE)
        self.assertEqual(rule.operator, ComparisonOperator.GREATER_THAN)

    def test_style_hash_calculation(self):
        """测试样式哈希计算"""
        style1 = {'bold': True, 'color': '#ff0000'}
        style2 = {'bold': True, 'color': '#ff0000'}
        style3 = {'bold': False, 'color': '#ff0000'}
        
        hash1 = self.style_manager._calculate_style_hash(style1)
        hash2 = self.style_manager._calculate_style_hash(style2)
        hash3 = self.style_manager._calculate_style_hash(style3)
        
        # 相同样式应该有相同哈希
        self.assertEqual(hash1, hash2)
        
        # 不同样式应该有不同哈希
        self.assertNotEqual(hash1, hash3)

    def test_css_property_conversion(self):
        """测试CSS属性转换"""
        style = {
            'bold': True,
            'italic': True,
            'font_size': 14,
            'font_color': '#ff0000',
            'bg_color': '#ffff00',
            'align': 'center'
        }
        
        css_props = self.style_manager._convert_style_to_css_dict(style)
        
        # 检查转换结果
        self.assertEqual(css_props['font-weight'], 'bold')
        self.assertEqual(css_props['font-style'], 'italic')
        self.assertEqual(css_props['font-size'], '14pt')
        self.assertEqual(css_props['color'], '#ff0000')
        self.assertEqual(css_props['background-color'], '#ffff00')
        self.assertEqual(css_props['text-align'], 'center')

    def test_color_interpolation(self):
        """测试颜色插值"""
        # 测试红到绿的插值
        red = '#ff0000'
        green = '#00ff00'
        
        # 中间值应该是混合色
        middle = self.style_manager._interpolate_color(red, green, 0.5)
        self.assertIsInstance(middle, str)
        self.assertTrue(middle.startswith('#'))
        self.assertEqual(len(middle), 7)
        
        # 0比例应该返回第一个颜色
        start = self.style_manager._interpolate_color(red, green, 0.0)
        self.assertEqual(start.lower(), red.lower())
        
        # 1比例应该返回第二个颜色
        end = self.style_manager._interpolate_color(red, green, 1.0)
        self.assertEqual(end.lower(), green.lower())

    def test_numeric_detection(self):
        """测试数值检测"""
        # 正确的数值
        self.assertTrue(self.style_manager._is_numeric('123'))
        self.assertTrue(self.style_manager._is_numeric('123.45'))
        self.assertTrue(self.style_manager._is_numeric('-123'))
        self.assertTrue(self.style_manager._is_numeric(123))
        self.assertTrue(self.style_manager._is_numeric(123.45))
        
        # 非数值
        self.assertFalse(self.style_manager._is_numeric('abc'))
        self.assertFalse(self.style_manager._is_numeric(''))
        self.assertFalse(self.style_manager._is_numeric(None))
        self.assertFalse(self.style_manager._is_numeric('123abc'))

    def test_rule_management(self):
        """测试规则管理"""
        # 添加规则
        rule = ConditionalRule(
            name="测试规则",
            type=ConditionalType.VALUE_RANGE,
            operator=ComparisonOperator.GREATER_THAN,
            values=[100],
            styles={'color': '#ff0000'},
            priority=10
        )
        
        initial_count = len(self.style_manager.conditional_rules)
        self.style_manager.add_conditional_rule(rule)
        
        # 检查规则已添加
        self.assertEqual(len(self.style_manager.conditional_rules), initial_count + 1)
        
        # 移除规则
        self.assertTrue(self.style_manager.remove_conditional_rule("测试规则"))
        self.assertEqual(len(self.style_manager.conditional_rules), initial_count)
        
        # 移除不存在的规则
        self.assertFalse(self.style_manager.remove_conditional_rule("不存在的规则"))

    def test_statistics_update(self):
        """测试统计信息更新"""
        # 初始统计
        initial_stats = self.style_manager.get_style_statistics()
        self.assertIn('total_styles', initial_stats)
        
        # 生成CSS类后检查统计
        self.style_manager.generate_css_classes(self.sample_sheet_data)
        updated_stats = self.style_manager.get_style_statistics()
        
        # 统计应该已更新
        self.assertGreaterEqual(updated_stats['total_styles'], 0)
        self.assertGreaterEqual(updated_stats['unique_styles'], 0)
        self.assertGreaterEqual(updated_stats['class_reuse_rate'], 0)


if __name__ == '__main__':
    unittest.main() 