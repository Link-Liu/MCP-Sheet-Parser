# html_converter.py
# HTML转换与样式映射

import logging
from typing import Any

from .config import Config, THEMES
from .data_validator import sanitize_for_html
from .style_manager import StyleManager, ConditionalRule, ConditionalType, ComparisonOperator


class HTMLConverter:
    """HTML转换器"""
    
    def __init__(self, sheet_data, config=None, theme='default', 
                 use_css_classes=True, conditional_rules=None):
        self.sheet_data = sheet_data
        self.config = config or Config()
        self.theme = theme
        self.logger = logging.getLogger(__name__)
        
        # 样式管理器
        self.style_manager = StyleManager(config)
        self.use_css_classes = use_css_classes
        
        # CSS类映射和条件样式
        self.css_class_map = {}
        self.conditional_styles = {}
        
        # 验证主题
        if theme not in THEMES:
            self.logger.warning(f"未知主题 '{theme}'，使用默认主题")
            self.theme = 'default'
        
        # 设置条件格式化规则
        if conditional_rules:
            for rule in conditional_rules:
                self.style_manager.add_conditional_rule(rule)

    def to_html(self, table_only=False, style_options=None):
        """
        转换为HTML
        
        Args:
            table_only: 是否只输出表格部分
            style_options: 样式选项
                - use_css_classes: 使用CSS类
                - semantic_names: 语义化类名
                - min_usage_threshold: 最小使用阈值
                - apply_conditional: 应用条件格式化
                - template: 样式模板名称
            
        Returns:
            HTML字符串
        """
        try:
            # 合并样式选项
            options = {
                'use_css_classes': self.use_css_classes,
                'semantic_names': True,
                'min_usage_threshold': 2,
                'apply_conditional': True,
                'template': None
            }
            if style_options:
                options.update(style_options)
            
            # 设置样式模板
            if options.get('template'):
                self.style_manager.set_template(options['template'])
            
            # 生成CSS类（如果启用）
            if options['use_css_classes']:
                css_content, self.css_class_map = self.style_manager.generate_css_classes(
                    self.sheet_data,
                    use_semantic_names=options['semantic_names'],
                    min_usage_threshold=options['min_usage_threshold']
                )
            else:
                css_content = ""
                self.css_class_map = {}
            
            # 应用条件格式化
            if options['apply_conditional']:
                self.conditional_styles = self.style_manager.apply_conditional_formatting(
                    self.sheet_data
                )
            
            if table_only:
                return self._generate_table_html()
            else:
                return self._generate_full_html(css_content)
                
        except Exception as e:
            self.logger.error(f"HTML转换失败: {e}")
            raise

    def _generate_full_html(self, css_content=""):
        """生成完整的HTML文档"""
        theme_config = THEMES[self.theme]
        
        html_parts = [
            '<!DOCTYPE html>',
            '<html lang="zh-CN">',
            '<head>',
            f'    <meta charset="{self.config.HTML_DEFAULT_ENCODING}">',
            '    <meta name="viewport" content="width=device-width, initial-scale=1.0">',
            f'    <title>{self.sheet_data.get("sheet_name", "工作表")}</title>',
            '    <style>',
            f'        body {{ {theme_config["body_style"]} }}',
            f'        table {{ {theme_config["table_style"]} }}',
            f'        td, th {{ {theme_config["cell_style"]} }}',
            f'        .header {{ {theme_config["header_style"]} }}',
            self._generate_responsive_css(),
            self._generate_comment_css(),
        ]
        
        # 添加生成的CSS类
        if css_content:
            html_parts.extend([
                '        /* 生成的CSS类 */',
                css_content
            ])
        
        html_parts.extend([
            '    </style>',
            '</head>',
            '<body>',
            self._generate_table_html(),
            self._generate_charts_html(),
            self._generate_comment_script(),
            '</body>',
            '</html>'
        ])
        
        return '\n'.join(html_parts)

    def _generate_table_html(self):
        """生成表格HTML"""
        data = self.sheet_data.get('data', [])
        styles = self.sheet_data.get('styles', [])
        merged_cells = self.sheet_data.get('merged_cells', [])
        comments = self.sheet_data.get('comments', {})
        hyperlinks = self.sheet_data.get('hyperlinks', {})
        
        if not data:
            return '<p>表格为空</p>'
        
        # 创建合并单元格映射
        merged_map = self._create_merged_map(merged_cells)
        
        # 生成表格
        table_classes = self._get_table_classes()
        table_parts = [
            f'<table {table_classes} border="{self.config.HTML_TABLE_BORDER}" ',
            f'cellspacing="{self.config.HTML_CELL_SPACING}" ',
            f'cellpadding="{self.config.HTML_CELL_PADDING}">'
        ]
        
        for row_idx, row in enumerate(data):
            row_classes = self._get_row_classes(row_idx)
            table_parts.append(f'    <tr{row_classes}>')
            
            for col_idx, cell_value in enumerate(row):
                # 检查是否被合并单元格覆盖
                if (row_idx, col_idx) in merged_map:
                    continue
                
                # 检查是否为合并单元格起始位置
                rowspan, colspan = self._get_span(row_idx, col_idx, merged_cells)
                
                # 清理单元格值
                safe_value = sanitize_for_html(cell_value)
                
                # 获取单元格样式
                cell_style = {}
                if row_idx < len(styles) and col_idx < len(styles[row_idx]):
                    cell_style = styles[row_idx][col_idx]
                
                # 生成单元格
                cell_html = self._generate_cell_html(
                    safe_value, row_idx, col_idx, rowspan, colspan, 
                    cell_style, comments, hyperlinks
                )
                
                table_parts.append(f'        {cell_html}')
            
            table_parts.append('    </tr>')
        
        table_parts.append('</table>')
        
        return '\n'.join(table_parts)

    def _get_table_classes(self):
        """获取表格CSS类"""
        classes = []
        
        # 添加模板类
        if (self.style_manager.current_template and 
            hasattr(self.style_manager.current_template, 'classes')):
            template_name = self.style_manager.current_template.name
            classes.append(f'{template_name}-table')
        
        return f'class="{" ".join(classes)}"' if classes else ''

    def _get_row_classes(self, row_idx):
        """获取行CSS类"""
        classes = []
        
        # 添加模板行类
        if (self.style_manager.current_template and 
            hasattr(self.style_manager.current_template, 'classes')):
            template_name = self.style_manager.current_template.name
            if row_idx == 0:
                classes.append(f'{template_name}-header')
            else:
                classes.append(f'{template_name}-row')
        
        return f' class="{" ".join(classes)}"' if classes else ''

    def _create_merged_map(self, merged_cells):
        """创建合并单元格映射"""
        merged_map = set()
        for start_row, start_col, end_row, end_col in merged_cells:
            for r in range(start_row, end_row + 1):
                for c in range(start_col, end_col + 1):
                    if r != start_row or c != start_col:
                        merged_map.add((r, c))
        return merged_map

    def _get_span(self, row_idx, col_idx, merged_cells):
        """获取单元格跨度"""
        for start_row, start_col, end_row, end_col in merged_cells:
            if row_idx == start_row and col_idx == start_col:
                rowspan = end_row - start_row + 1
                colspan = end_col - start_col + 1
                return rowspan, colspan
        return 1, 1

    def _generate_cell_html(self, value, row_idx, col_idx, rowspan, colspan, cell_style, comments, hyperlinks):
        """生成单元格HTML"""
        # 基本单元格内容
        cell_content = value if value else '&nbsp;'
        
        # 检查是否为公式单元格
        is_formula = self._is_formula_cell(value)
        formula_info = None
        
        if is_formula and hasattr(self, 'sheet_data'):
            formulas = self.sheet_data.get('formulas', {})
            formula_key = f"{row_idx}_{col_idx}"
            formula_info = formulas.get(formula_key)
        
        # 处理公式显示
        if formula_info:
            cell_content = self._format_formula_cell(formula_info, value)
        
        # 处理超链接
        link_key = f"{row_idx}_{col_idx}"
        if link_key in hyperlinks and self.config.INCLUDE_HYPERLINKS:
            href = hyperlinks[link_key]
            if not is_formula:  # 公式单元格不包装超链接
                cell_content = f'<a href="{href}" target="_blank">{cell_content}</a>'
        
        # 处理注释
        comment_html = ''
        if link_key in comments and self.config.INCLUDE_COMMENTS:
            comment_text = sanitize_for_html(comments[link_key])
            comment_html = f' title="{comment_text}"'
        
        # 处理公式悬停提示
        if formula_info and self.config.SHOW_FORMULA_TEXT:
            formula_text = formula_info.original_formula
            description = formula_info.description
            tooltip = f"公式: {formula_text}"
            if description:
                tooltip += f" ({description})"
            if formula_info.error:
                tooltip += f" [错误: {formula_info.error.value}]"
            comment_html = f' title="{sanitize_for_html(tooltip)}" data-formula="{sanitize_for_html(formula_text)}"'
        
        # 构建属性
        attrs = []
        if rowspan > 1:
            attrs.append(f'rowspan="{rowspan}"')
        if colspan > 1:
            attrs.append(f'colspan="{colspan}"')
        
        # 处理样式和类
        classes = []
        inline_styles = []
        
        # 添加公式类
        if formula_info:
            classes.append('formula-cell')
            if formula_info.error:
                classes.append('formula-error')
        
        # 获取CSS类（如果启用）
        if self.use_css_classes and (row_idx, col_idx) in self.css_class_map:
            classes.append(self.css_class_map[(row_idx, col_idx)])
        else:
            # 使用内联样式
            style_css = self._convert_style_to_css(cell_style)
            if style_css:
                inline_styles.append(style_css)
        
        # 添加条件格式化样式
        if (row_idx, col_idx) in self.conditional_styles:
            conditional_css = self.conditional_styles[(row_idx, col_idx)]
            inline_styles.append(conditional_css)
        
        # 添加类属性
        if classes:
            attrs.append(f'class="{" ".join(classes)}"')
        
        # 添加内联样式
        if inline_styles:
            combined_style = "; ".join(inline_styles)
            attrs.append(f'style="{combined_style}"')
        
        attrs_str = ' ' + ' '.join(attrs) if attrs else ''
        
        # 第一行作为表头
        tag = 'th' if row_idx == 0 else 'td'
        
        return f'<{tag}{attrs_str}{comment_html}>{cell_content}</{tag}>'

    def _is_formula_cell(self, value: Any) -> bool:
        """判断是否为公式单元格"""
        return isinstance(value, str) and value.startswith('=')

    def _format_formula_cell(self, formula_info, original_value: str) -> str:
        """格式化公式单元格内容"""
        if not formula_info:
            return original_value
        
        parts = []
        
        # 添加公式指示符
        parts.append('<span class="formula-indicator">ƒ</span>')
        
        # 显示计算结果或错误
        if formula_info.error:
            if self.config.SHOW_FORMULA_ERRORS:
                parts.append(f'<span class="formula-result">{formula_info.error.value}</span>')
            else:
                parts.append('<span class="formula-result">#ERROR</span>')
        elif formula_info.calculated_value is not None:
            # 格式化计算结果
            result = formula_info.calculated_value
            if isinstance(result, float):
                # 保留合理的小数位数
                if result.is_integer():
                    result_str = str(int(result))
                else:
                    result_str = f"{result:.6g}"  # 最多6位有效数字
            else:
                result_str = str(result)
            
            parts.append(f'<span class="formula-result">{sanitize_for_html(result_str)}</span>')
        else:
            # 显示原始公式文本
            parts.append(f'<span class="formula-text">{sanitize_for_html(original_value)}</span>')
        
        return ''.join(parts)

    def _convert_style_to_css(self, style):
        """将样式字典转换为CSS字符串"""
        if not style:
            return ''
        
        css_parts = []
        
        # 字体样式
        if style.get('bold'):
            css_parts.append('font-weight: bold')
        if style.get('italic'):
            css_parts.append('font-style: italic')
        if style.get('underline') == 'single':
            css_parts.append('text-decoration: underline')
        if style.get('strike'):
            css_parts.append('text-decoration: line-through')
        
        # 字体大小
        if style.get('font_size'):
            css_parts.append(f'font-size: {style["font_size"]}pt')
        
        # 字体名称
        if style.get('font_name'):
            css_parts.append(f'font-family: "{style["font_name"]}"')
        
        # 字体颜色
        if style.get('font_color'):
            css_parts.append(f'color: {style["font_color"]}')
        
        # 背景颜色
        if style.get('bg_color'):
            css_parts.append(f'background-color: {style["bg_color"]}')
        
        # 文本对齐
        if style.get('align'):
            css_parts.append(f'text-align: {style["align"]}')
        if style.get('valign'):
            valign_map = {
                'top': 'vertical-align: top',
                'center': 'vertical-align: middle',
                'bottom': 'vertical-align: bottom'
            }
            if style['valign'] in valign_map:
                css_parts.append(valign_map[style['valign']])
        
        # 边框样式
        if style.get('border'):
            border_css = self._convert_border_to_css(style['border'])
            if border_css:
                css_parts.extend(border_css)
        
        # 文本换行
        if style.get('wrap_text'):
            css_parts.append('white-space: pre-wrap')
        
        return '; '.join(css_parts)

    def _convert_border_to_css(self, border):
        """将边框样式转换为CSS"""
        css_parts = []
        
        for side in ['top', 'bottom', 'left', 'right']:
            if side in border:
                border_info = border[side]
                style_name = border_info.get('style', 'solid')
                color = border_info.get('color', '#000000')
                
                # 转换边框样式
                style_map = {
                    'thin': 'solid 1px',
                    'thick': 'solid 3px',
                    'medium': 'solid 2px',
                    'dashed': 'dashed 1px',
                    'dotted': 'dotted 1px',
                    'double': 'double 3px'
                }
                
                css_style = style_map.get(style_name, f'{style_name} 1px')
                css_parts.append(f'border-{side}: {css_style} {color}')
        
        return css_parts

    def _generate_responsive_css(self):
        """生成响应式CSS"""
        return """
        @media screen and (max-width: 768px) {
            table { font-size: 12px; }
            td, th { padding: 4px; }
        }"""

    def _generate_comment_css(self):
        """生成注释CSS"""
        return """
        td[title], th[title] {
            position: relative;
            cursor: help;
        }
        td[title]:after, th[title]:after {
            content: '💬';
            position: absolute;
            top: 2px;
            right: 2px;
            font-size: 10px;
            color: #ff6b6b;
        }
        
        /* 公式单元格样式 */
        .formula-cell {
            position: relative;
            background: linear-gradient(135deg, #f8f9ff 0%, #e6f3ff 100%);
            border-left: 3px solid #4472C4;
        }
        
        .formula-cell:hover {
            background: linear-gradient(135deg, #e6f3ff 0%, #cce7ff 100%);
        }
        
        .formula-indicator {
            position: absolute;
            top: 1px;
            left: 2px;
            font-size: 10px;
            color: #4472C4;
            font-weight: bold;
            pointer-events: none;
        }
        
        .formula-result {
            padding-left: 12px;
        }
        
        .formula-error {
            color: #dc3545;
            font-weight: bold;
            background-color: #f8d7da;
        }
        
        .formula-text {
            font-family: 'Courier New', monospace;
            font-size: 11px;
            color: #666;
            font-style: italic;
        }"""

    def _generate_charts_html(self):
        """生成图表HTML"""
        charts = self.sheet_data.get('charts', [])
        if not charts:
            return ''
        
        html_parts = [
            '<div class="charts-container" style="margin-top: 30px;">',
            '    <h2 style="color: #333; border-bottom: 2px solid #4472C4; padding-bottom: 10px;">📊 数据图表</h2>'
        ]
        
        for i, chart in enumerate(charts):
            chart_title = chart.get('title', f'图表 {i+1}')
            chart_svg = chart.get('svg', '')
            chart_type = chart.get('type', 'unknown')
            
            html_parts.extend([
                f'    <div class="chart-item" style="margin: 20px 0; padding: 20px; background: #f8f9fa; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">',
                f'        <h3 style="margin-top: 0; color: #495057; text-align: center;">{chart_title}</h3>',
                f'        <div class="chart-content" style="text-align: center; overflow-x: auto;">',
                f'            {chart_svg}',
                f'        </div>',
                f'        <p style="text-align: center; color: #6c757d; font-size: 12px; margin-bottom: 0;">类型: {chart_type}</p>',
                f'    </div>'
            ])
        
        html_parts.append('</div>')
        return '\n'.join(html_parts)

    def _generate_comment_script(self):
        """生成注释脚本"""
        return """
    <script>
        // 为带有title属性的单元格添加交互效果
        document.addEventListener('DOMContentLoaded', function() {
            const cells = document.querySelectorAll('td[title], th[title]');
            cells.forEach(cell => {
                cell.addEventListener('mouseenter', function() {
                    this.style.backgroundColor = '#f0f8ff';
                });
                cell.addEventListener('mouseleave', function() {
                    this.style.backgroundColor = '';
                });
            });
            
            // 图表响应式处理
            const charts = document.querySelectorAll('.chart-content svg');
            charts.forEach(chart => {
                chart.style.maxWidth = '100%';
                chart.style.height = 'auto';
            });
        });
    </script>"""

    def to_html_table_only(self):
        """只生成表格HTML（向后兼容）"""
        return self.to_html(table_only=True)

    def export_to_file(self, file_path, table_only=False, style_options=None):
        """导出HTML到文件"""
        from .utils import ensure_output_dir
        from .security import validate_output_path
        
        # 验证输出路径
        is_safe, warnings = validate_output_path(file_path)
        if not is_safe:
            raise ValueError(f"输出路径不安全: {'; '.join(warnings)}")
        
        # 确保输出目录存在
        ensure_output_dir(file_path)
        
        # 生成HTML内容
        html_content = self.to_html(table_only, style_options)
        
        # 写入文件
        try:
            with open(file_path, 'w', encoding=self.config.HTML_DEFAULT_ENCODING) as f:
                f.write(html_content)
            
            self.logger.info(f"HTML已导出到: {file_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"导出HTML失败: {e}")
            raise

    def get_style_statistics(self):
        """获取样式统计信息"""
        return self.style_manager.get_style_statistics()

    def add_conditional_rule(self, rule: ConditionalRule):
        """添加条件格式化规则"""
        self.style_manager.add_conditional_rule(rule)

    def remove_conditional_rule(self, rule_name: str):
        """移除条件格式化规则"""
        return self.style_manager.remove_conditional_rule(rule_name)

    def set_style_template(self, template_name: str):
        """设置样式模板"""
        return self.style_manager.set_template(template_name)
