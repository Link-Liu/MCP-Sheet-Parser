#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
高级样式控制演示程序
展示CSS类生成、条件格式化和样式模板功能
"""

import os
import sys
from pathlib import Path

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from mcp_sheet_parser.parser import SheetParser
from mcp_sheet_parser.html_converter import HTMLConverter
from mcp_sheet_parser.style_manager import (
    StyleManager, ConditionalRule, ConditionalType, ComparisonOperator
)
from mcp_sheet_parser.config import Config


def create_sample_financial_data():
    """创建示例财务数据"""
    return {
        'sheet_name': '财务报表示例',
        'data': [
            ['项目', '预算', '实际', '差异', '状态'],
            ['营销费用', '10000', '8500', '1500', '节约'],
            ['研发投入', '20000', '22000', '-2000', '超支'],
            ['办公费用', '5000', '4800', '200', '节约'],
            ['人力成本', '30000', '31500', '-1500', '超支'],
            ['设备采购', '15000', '12000', '3000', '节约'],
            ['培训费用', '3000', '2800', '200', '节约'],
            ['差旅费用', '8000', '9200', '-1200', '超支']
        ],
        'styles': [
            # 表头样式
            [
                {'bold': True, 'bg_color': '#4472C4', 'font_color': '#ffffff'},
                {'bold': True, 'bg_color': '#4472C4', 'font_color': '#ffffff'},
                {'bold': True, 'bg_color': '#4472C4', 'font_color': '#ffffff'},
                {'bold': True, 'bg_color': '#4472C4', 'font_color': '#ffffff'},
                {'bold': True, 'bg_color': '#4472C4', 'font_color': '#ffffff'}
            ],
            # 数据行样式
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {}]
        ],
        'merged_cells': [],
        'comments': {
            '1_3': '差异 = 预算 - 实际',
            '2_4': '超支需要审批'
        },
        'hyperlinks': {}
    }


def create_sample_analytics_data():
    """创建示例分析数据"""
    return {
        'sheet_name': '销售分析示例',
        'data': [
            ['销售员', '一月', '二月', '三月', '总计', '完成率'],
            ['张三', '85000', '92000', '88000', '265000', '106%'],
            ['李四', '78000', '85000', '79000', '242000', '97%'],
            ['王五', '95000', '89000', '91000', '275000', '110%'],
            ['赵六', '72000', '68000', '74000', '214000', '86%'],
            ['钱七', '88000', '94000', '96000', '278000', '111%'],
            ['孙八', '65000', '71000', '69000', '205000', '82%'],
            ['周九', '91000', '87000', '93000', '271000', '108%']
        ],
        'styles': [
            # 表头样式
            [
                {'bold': True, 'bg_color': '#70AD47', 'font_color': '#ffffff'},
                {'bold': True, 'bg_color': '#70AD47', 'font_color': '#ffffff', 'align': 'center'},
                {'bold': True, 'bg_color': '#70AD47', 'font_color': '#ffffff', 'align': 'center'},
                {'bold': True, 'bg_color': '#70AD47', 'font_color': '#ffffff', 'align': 'center'},
                {'bold': True, 'bg_color': '#70AD47', 'font_color': '#ffffff', 'align': 'center'},
                {'bold': True, 'bg_color': '#70AD47', 'font_color': '#ffffff', 'align': 'center'}
            ],
            # 数据行样式（数值右对齐）
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {'align': 'right', 'bold': True}, {'align': 'center'}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {'align': 'right', 'bold': True}, {'align': 'center'}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {'align': 'right', 'bold': True}, {'align': 'center'}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {'align': 'right', 'bold': True}, {'align': 'center'}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {'align': 'right', 'bold': True}, {'align': 'center'}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {'align': 'right', 'bold': True}, {'align': 'center'}],
            [{}, {'align': 'right'}, {'align': 'right'}, {'align': 'right'}, {'align': 'right', 'bold': True}, {'align': 'center'}]
        ],
        'merged_cells': [],
        'comments': {
            '0_5': '完成率 = 实际销售 / 目标销售 * 100%'
        },
        'hyperlinks': {}
    }


def demo_css_class_generation():
    """演示CSS类生成功能"""
    print("🎨 === CSS类生成演示 ===")
    
    # 创建示例数据
    sheet_data = create_sample_financial_data()
    
    # 创建样式管理器
    style_manager = StyleManager()
    
    # 生成CSS类
    print("生成CSS类...")
    css_content, cell_class_map = style_manager.generate_css_classes(
        sheet_data,
        use_semantic_names=True,
        min_usage_threshold=1
    )
    
    print(f"✅ 生成了 {len(style_manager.style_classes)} 个CSS类")
    
    # 显示生成的CSS类
    print("\n生成的CSS类:")
    for class_name, style_class in style_manager.style_classes.items():
        print(f"  .{class_name} - {style_class.description} (使用{style_class.usage_count}次)")
    
    # 显示CSS内容片段
    print(f"\nCSS内容预览 (前500字符):")
    print(css_content[:500] + "..." if len(css_content) > 500 else css_content)
    
    # 显示统计信息
    stats = style_manager.get_style_statistics()
    print(f"\n📊 统计信息:")
    print(f"  总样式数: {stats['total_styles']}")
    print(f"  独特样式数: {stats['unique_styles']}")
    print(f"  CSS类复用率: {stats['class_reuse_rate']:.1f}%")
    
    return style_manager, css_content, cell_class_map


def demo_conditional_formatting():
    """演示条件格式化功能"""
    print("\n🌈 === 条件格式化演示 ===")
    
    # 创建示例数据
    sheet_data = create_sample_financial_data()
    
    # 创建样式管理器
    style_manager = StyleManager()
    
    # 添加财务条件格式化规则
    print("添加条件格式化规则...")
    
    # 1. 正负值规则
    positive_rule = ConditionalRule(
        name="正值高亮",
        type=ConditionalType.VALUE_RANGE,
        operator=ComparisonOperator.GREATER_THAN,
        values=[0],
        styles={'color': '#28a745', 'font-weight': 'bold'},
        priority=10,
        description="突出显示正值"
    )
    style_manager.add_conditional_rule(positive_rule)
    
    negative_rule = ConditionalRule(
        name="负值警告",
        type=ConditionalType.VALUE_RANGE,
        operator=ComparisonOperator.LESS_THAN,
        values=[0],
        styles={'color': '#dc3545', 'font-weight': 'bold', 'background-color': '#f8d7da'},
        priority=10,
        description="警告负值"
    )
    style_manager.add_conditional_rule(negative_rule)
    
    # 2. 数据条规则
    data_bars_rule = ConditionalRule(
        name="预算数据条",
        type=ConditionalType.DATA_BARS,
        operator=ComparisonOperator.GREATER_THAN,
        values=['#4472C4'],
        styles={},
        priority=5,
        description="预算金额数据条"
    )
    style_manager.add_conditional_rule(data_bars_rule)
    
    # 3. 重复值规则  
    duplicate_rule = ConditionalRule(
        name="重复状态",
        type=ConditionalType.DUPLICATE_VALUES,
        operator=ComparisonOperator.EQUAL,
        values=[],
        styles={'background-color': '#fff3cd', 'color': '#856404'},
        priority=8,
        description="标记重复状态"
    )
    style_manager.add_conditional_rule(duplicate_rule)
    
    # 应用条件格式化
    conditional_styles = style_manager.apply_conditional_formatting(sheet_data)
    
    print(f"✅ 应用了 {len(conditional_styles)} 个条件样式")
    
    # 显示应用的条件样式
    print("\n应用的条件样式:")
    for (row, col), style in conditional_styles.items():
        cell_value = sheet_data['data'][row][col] if row < len(sheet_data['data']) and col < len(sheet_data['data'][row]) else ''
        print(f"  单元格[{row},{col}] '{cell_value}': {style}")
    
    # 显示统计信息
    stats = style_manager.get_style_statistics()
    print(f"\n📊 条件格式化统计:")
    print(f"  应用条件规则: {stats.get('conditional_rules_applied', 0)}次")
    
    return style_manager, conditional_styles


def demo_style_templates():
    """演示样式模板功能"""
    print("\n📋 === 样式模板演示 ===")
    
    # 创建样式管理器
    style_manager = StyleManager()
    
    # 显示可用模板
    print("可用的样式模板:")
    for name, template in style_manager.templates.items():
        print(f"  {name}: {template.description}")
    
    # 测试每个模板
    templates_to_test = ['business', 'financial', 'analytics']
    
    for template_name in templates_to_test:
        print(f"\n🎨 测试模板: {template_name}")
        
        # 设置模板
        success = style_manager.set_template(template_name)
        if success:
            print(f"✅ 成功应用模板: {template_name}")
            
            # 显示模板的全局CSS
            if style_manager.current_template.global_css:
                css_preview = style_manager.current_template.global_css[:200]
                print(f"  全局CSS预览: {css_preview}...")
            
        else:
            print(f"❌ 应用模板失败: {template_name}")
    
    return style_manager


def demo_complete_workflow():
    """演示完整的高级样式控制工作流"""
    print("\n🚀 === 完整工作流演示 ===")
    
    # 创建输出目录
    output_dir = Path("demo_output")
    output_dir.mkdir(exist_ok=True)
    
    # 准备两套数据
    datasets = [
        ("financial", create_sample_financial_data()),
        ("analytics", create_sample_analytics_data())
    ]
    
    for data_type, sheet_data in datasets:
        print(f"\n📊 处理 {data_type} 数据...")
        
        # 创建配置
        config = Config()
        
        # 创建HTML转换器，启用高级样式功能
        converter = HTMLConverter(
            sheet_data,
            config=config,
            theme='default',
            use_css_classes=True
        )
        
        # 设置样式模板
        template_name = 'financial' if data_type == 'financial' else 'analytics'
        converter.set_style_template(template_name)
        print(f"  应用样式模板: {template_name}")
        
        # 添加条件格式化规则
        if data_type == 'financial':
            # 财务数据的条件规则
            positive_rule = ConditionalRule(
                name="正值",
                type=ConditionalType.VALUE_RANGE,
                operator=ComparisonOperator.GREATER_THAN,
                values=[0],
                styles={'color': '#28a745', 'font-weight': 'bold'},
                priority=10
            )
            converter.add_conditional_rule(positive_rule)
            
            negative_rule = ConditionalRule(
                name="负值",
                type=ConditionalType.VALUE_RANGE,
                operator=ComparisonOperator.LESS_THAN,
                values=[0],
                styles={'color': '#dc3545', 'background-color': '#f8d7da'},
                priority=10
            )
            converter.add_conditional_rule(negative_rule)
            
        else:  # analytics
            # 分析数据的条件规则
            high_performance = ConditionalRule(
                name="高绩效",
                type=ConditionalType.VALUE_RANGE,
                operator=ComparisonOperator.GREATER_THAN,
                values=[100000],
                styles={'background-color': '#d4edda', 'font-weight': 'bold'},
                priority=10
            )
            converter.add_conditional_rule(high_performance)
            
            data_bars = ConditionalRule(
                name="销售数据条",
                type=ConditionalType.DATA_BARS,
                operator=ComparisonOperator.GREATER_THAN,
                values=['#70AD47'],
                styles={},
                priority=5
            )
            converter.add_conditional_rule(data_bars)
        
        # 配置样式选项
        style_options = {
            'use_css_classes': True,
            'semantic_names': True,
            'min_usage_threshold': 1,
            'apply_conditional': True,
            'template': template_name
        }
        
        # 生成HTML
        html_content = converter.to_html(style_options=style_options)
        
        # 保存文件
        output_file = output_dir / f"{data_type}_advanced_styles.html"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"  ✅ 保存到: {output_file}")
        
        # 显示统计信息
        stats = converter.get_style_statistics()
        print(f"  📈 样式统计:")
        print(f"    - 总样式数: {stats.get('total_styles', 0)}")
        print(f"    - 独特样式数: {stats.get('unique_styles', 0)}")
        print(f"    - CSS类复用率: {stats.get('class_reuse_rate', 0):.1f}%")
        print(f"    - 条件格式化应用: {stats.get('conditional_rules_applied', 0)}次")
    
    print(f"\n🎉 演示完成！输出文件保存在 {output_dir} 目录中")
    return output_dir


def main():
    """主演示函数"""
    print("🎨 高级样式控制功能演示")
    print("=" * 50)
    
    try:
        # 1. CSS类生成演示
        demo_css_class_generation()
        
        # 2. 条件格式化演示
        demo_conditional_formatting()
        
        # 3. 样式模板演示
        demo_style_templates()
        
        # 4. 完整工作流演示
        output_dir = demo_complete_workflow()
        
        print(f"\n✨ 所有演示完成！")
        print(f"📁 生成的文件位于: {output_dir.absolute()}")
        print(f"🌐 您可以在浏览器中打开生成的HTML文件查看效果")
        
        return 0
        
    except Exception as e:
        print(f"\n❌ 演示过程中出现错误: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit(main()) 