# test_alignment_demo.py
# 演示增强的对齐方式支持

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..'))

from mcp_sheet_parser.config.style import StyleConfig
from mcp_sheet_parser.html_converter import HTMLConverter
from mcp_sheet_parser.style_manager import StyleManager
from mcp_sheet_parser.config import Config


def demo_alignment_support():
    """演示对齐方式支持"""
    print("=" * 60)
    print("MCP-Sheet-Parser 增强对齐方式支持演示")
    print("=" * 60)
    
    # 初始化配置
    config = Config()
    style_config = StyleConfig()
    html_converter = HTMLConverter(config=config)
    style_manager = StyleManager(config=config)
    
    print("\n1. 中文对齐方式支持")
    print("-" * 30)
    
    # 测试中文水平对齐
    chinese_horizontal = ['左对齐', '居中', '右对齐', '两端对齐', '分散对齐', '填充', '常规']
    for alignment in chinese_horizontal:
        result = style_config.get_alignment_style(alignment)
        print(f"  {alignment:8} -> {result}")
    
    # 测试中文垂直对齐
    chinese_vertical = ['顶端对齐', '垂直居中', '底端对齐', '垂直两端对齐', '垂直分散对齐']
    for alignment in chinese_vertical:
        result = style_config.get_vertical_alignment_style(alignment)
        print(f"  {alignment:10} -> {result}")
    
    print("\n2. 英文别名支持")
    print("-" * 30)
    
    # 测试英文别名
    english_aliases = [
        ('start', 'left'),
        ('end', 'right'),
        ('middle', 'center'),
        ('justified', 'justify'),
        ('distribute', 'justify')
    ]
    
    for alias, expected in english_aliases:
        result = style_config.get_alignment_style(alias)
        print(f"  {alias:10} -> {result} (期望: {expected})")
    
    print("\n3. 数字代码支持")
    print("-" * 30)
    
    # 测试数字代码
    numeric_codes = [
        ('0', 'general'),
        ('1', 'left'),
        ('2', 'center'),
        ('3', 'right'),
        ('4', 'fill'),
        ('5', 'justify'),
        ('6', 'centerContinuous'),
        ('7', 'distributed')
    ]
    
    for code, expected in numeric_codes:
        result = style_config.get_alignment_style(code)
        print(f"  {code} -> {result} (期望: {expected})")
    
    print("\n4. 大小写不敏感支持")
    print("-" * 30)
    
    # 测试大小写不敏感
    case_tests = ['LEFT', 'Center', 'RIGHT', 'JUSTIFY']
    for test in case_tests:
        result = style_config.get_alignment_style(test)
        print(f"  {test:8} -> {result}")
    
    print("\n5. 验证功能")
    print("-" * 30)
    
    # 测试验证功能
    valid_alignments = ['left', 'center', '右对齐', '1']
    invalid_alignments = ['invalid', None, '']
    
    print("有效对齐方式:")
    for alignment in valid_alignments:
        is_valid = style_config.is_valid_alignment(alignment)
        print(f"  {alignment:10} -> {is_valid}")
    
    print("无效对齐方式:")
    for alignment in invalid_alignments:
        is_valid = style_config.is_valid_alignment(alignment)
        print(f"  {alignment:10} -> {is_valid}")
    
    print("\n6. 支持的对齐方式列表")
    print("-" * 30)
    
    horizontal_alignments = style_config.get_supported_alignments()
    vertical_alignments = style_config.get_supported_vertical_alignments()
    
    print(f"水平对齐方式: {horizontal_alignments}")
    print(f"垂直对齐方式: {vertical_alignments}")
    
    print("\n7. HTML转换器测试")
    print("-" * 30)
    
    # 测试HTML转换器
    test_styles = [
        {'align': '居中', 'valign': '垂直居中'},
        {'align': '右对齐', 'valign': '底端对齐'},
        {'align': '2', 'valign': '1'},  # 数字代码
        {'align': '两端对齐', 'valign': '顶端对齐'}
    ]
    
    for i, style in enumerate(test_styles, 1):
        css_styles = html_converter._apply_cell_styles(style)
        print(f"测试 {i}: {style}")
        for css_style in css_styles:
            if 'text-align' in css_style or 'vertical-align' in css_style:
                print(f"  {css_style}")
    
    print("\n8. 样式管理器测试")
    print("-" * 30)
    
    # 测试样式管理器
    for i, style in enumerate(test_styles, 1):
        css_props = style_manager._convert_style_to_css_dict(style)
        print(f"测试 {i}: {style}")
        for prop, value in css_props.items():
            if 'text-align' in prop or 'vertical-align' in prop:
                print(f"  {prop}: {value}")
    
    print("\n9. 边界情况测试")
    print("-" * 30)
    
    # 测试边界情况
    edge_cases = [None, '', '   ', 'unknown']
    for case in edge_cases:
        horizontal = style_config.get_alignment_style(case)
        vertical = style_config.get_vertical_alignment_style(case)
        print(f"  {case:10} -> 水平: {horizontal}, 垂直: {vertical}")
    
    print("\n" + "=" * 60)
    print("演示完成！")
    print("=" * 60)


def demo_alignment_in_html():
    """演示在HTML中的对齐方式效果"""
    print("\n" + "=" * 60)
    print("HTML对齐方式效果演示")
    print("=" * 60)
    
    # 创建测试数据
    test_data = [
        {
            'sheet_name': '对齐方式演示',
            'rows': 4,
            'cols': 4,
            'data': [
                ['左对齐', '居中', '右对齐', '两端对齐'],
                ['顶端对齐', '垂直居中', '底端对齐', '组合对齐'],
                ['数字代码1', '数字代码2', '数字代码3', '数字代码4'],
                ['中文左对齐', '中文居中', '中文右对齐', '中文两端对齐']
            ],
            'styles': [
                [
                    {'align': 'left'}, {'align': 'center'}, {'align': 'right'}, {'align': 'justify'}
                ],
                [
                    {'valign': 'top'}, {'valign': 'middle'}, {'valign': 'bottom'}, 
                    {'align': 'center', 'valign': 'middle'}
                ],
                [
                    {'align': '1'}, {'align': '2'}, {'align': '3'}, {'align': '5'}
                ],
                [
                    {'align': '左对齐'}, {'align': '居中'}, {'align': '右对齐'}, {'align': '两端对齐'}
                ]
            ],
            'merged_cells': [],
            'comments': {},
            'hyperlinks': {}
        }
    ]
    
    # 转换HTML
    config = Config()
    html_converter = HTMLConverter(config=config)
    
    html_content = html_converter.convert_to_html(
        test_data, 
        'alignment_demo.html', 
        theme='default', 
        title='对齐方式演示'
    )
    
    # 保存HTML文件
    with open('demo/alignment_demo.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("HTML文件已生成: demo/alignment_demo.html")
    print("请在浏览器中打开查看对齐方式效果")


if __name__ == '__main__':
    demo_alignment_support()
    demo_alignment_in_html() 