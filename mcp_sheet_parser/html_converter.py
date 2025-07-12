# html_converter.py
# HTML转换器 - 将解析的表格数据转换为HTML格式

import os
import re
import logging
from typing import Dict, List, Any, Optional, Tuple
from .config import Config, THEMES
from .utils import setup_logger, get_file_extension
from .security import validate_file_path as security_validate_file_path, validate_output_path
from .formula_processor import FormulaInfo, FormulaError
from .exceptions import (
    ErrorHandler, ErrorContext, ErrorSeverity,
    HTMLConversionError, SecurityError, ConfigurationError,
    error_handler, safe_execute, create_error_context
)
import html


class HTMLConverter:
    def __init__(self, sheet_data=None, config=None, theme='default'):
        """
        HTML转换器初始化
        
        Args:
            sheet_data: 工作表数据（向后兼容）
            config: 配置对象
            theme: 主题名称（向后兼容）
        """
        self.config = config or Config()
        self.logger = setup_logger(__name__)
        self.error_handler = ErrorHandler(self.logger)
        
        # 向后兼容：保存单个工作表数据
        self.sheet_data = sheet_data
        self.theme = theme
        
    def to_html(self, table_only=False, title="表格数据") -> str:
        """
        将工作表数据转换为HTML（向后兼容方法）
        
        Args:
            table_only: 是否只输出表格部分
            title: HTML页面标题
            
        Returns:
            HTML字符串
        """
        if not self.sheet_data:
            return "<p>没有数据可转换</p>"
        
        # 如果只要表格，直接生成表格HTML
        if table_only:
            theme_config = self._validate_theme(self.theme, create_error_context("表格生成"))
            return self._convert_sheet_to_html(self.sheet_data, theme_config, True)
        
        # 生成完整HTML页面
        theme_config = self._validate_theme(self.theme, create_error_context("HTML生成"))
        return self._generate_html_content([self.sheet_data], theme_config, title, True, 
                                         create_error_context("HTML生成"))
    
    def export_to_file(self, output_path: str, title="表格数据") -> bool:
        """
        导出HTML到文件（向后兼容方法）
        
        Args:
            output_path: 输出文件路径
            title: HTML页面标题
            
        Returns:
            是否成功
        """
        try:
            if not self.sheet_data:
                return False
            
            html_content = self.to_html(table_only=False, title=title)
            
            context = create_error_context("文件导出", file_path=output_path)
            self._validate_output_path(output_path, context)
            self._write_html_file(output_path, html_content, context)
            
            return True
        except Exception as e:
            self.logger.error(f"导出HTML文件失败: {e}")
            return False
    
    def convert_to_html(self, sheets_data: List[Dict], output_path: str, 
                       theme: str = 'default', title: str = "表格数据",
                       include_styles: bool = True) -> str:
        """
        将表格数据转换为HTML格式
        
        Args:
            sheets_data: 解析的表格数据列表
            output_path: 输出文件路径
            theme: 主题名称
            title: HTML页面标题
            include_styles: 是否包含样式
            
        Returns:
            str: 输出文件路径
        """
        context = create_error_context(
            "HTML转换", 
            file_path=output_path,
            additional_info={'theme': theme, 'title': title}
        )
        
        try:
            self.logger.info(f"开始HTML转换，主题: {theme}")
            
            # 输出路径安全检查
            self._validate_output_path(output_path, context)
            
            # 验证主题配置
            theme_config = self._validate_theme(theme, context)
            
            # 生成HTML内容
            html_content = self._generate_html_content(
                sheets_data, theme_config, title, include_styles, context
            )
            
            # 写入文件
            self._write_html_file(output_path, html_content, context)
            
            self.logger.info(f"HTML转换完成: {output_path}")
            return output_path
            
        except Exception as e:
            custom_error = self.error_handler.handle_error(e, context)
            raise custom_error
    
    def _validate_output_path(self, output_path: str, context: ErrorContext):
        """验证输出路径安全性"""
        # 基本路径验证
        is_valid, path_warnings = validate_output_path(output_path)
        if not is_valid:
            raise HTMLConversionError(f"输出路径无效: {output_path}, 警告: {'; '.join(path_warnings)}", context=context)
        
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
                self.logger.info(f"创建输出目录: {output_dir}")
            except Exception as e:
                raise HTMLConversionError(
                    f"无法创建输出目录 {output_dir}: {e}", 
                    context=context, 
                    original_error=e
                )
    
    def _validate_theme(self, theme: str, context: ErrorContext) -> Dict:
        """验证并获取主题配置"""
        if theme not in THEMES:
            available_themes = ', '.join(THEMES.keys())
            raise ConfigurationError(
                'theme',
                f"不支持的主题 '{theme}'，可用主题: {available_themes}",
                context=context
            )
        
        theme_config = THEMES[theme]
        self.logger.info(f"使用主题: {theme_config['name']}")
        return theme_config
    
    def _generate_html_content(self, sheets_data: List[Dict], theme_config: Dict, 
                             title: str, include_styles: bool, context: ErrorContext) -> str:
        """生成HTML内容"""
        try:
            html_parts = []
            
            # HTML头部
            html_parts.append(self._generate_html_header(theme_config, title, include_styles))
            
            # 主体内容
            html_parts.append('<body>')
            html_parts.append(f'<h1>{title}</h1>')
            
            # 处理每个工作表
            for i, sheet in enumerate(sheets_data):
                sheet_context = create_error_context(
                    "工作表HTML转换",
                    sheet_name=sheet.get('sheet_name', f'Sheet{i+1}'),
                    additional_info={'sheet_index': i}
                )
                
                sheet_html = safe_execute(
                    self._convert_sheet_to_html,
                    sheet, theme_config, include_styles,
                    operation=f"工作表转换-{sheet.get('sheet_name', f'Sheet{i+1}')}",
                    default_value=f"<!-- 工作表 {sheet.get('sheet_name', f'Sheet{i+1}')} 转换失败 -->"
                )
                
                html_parts.append(sheet_html)
            
            html_parts.append('</body>')
            html_parts.append('</html>')
            
            return '\n'.join(html_parts)
            
        except Exception as e:
            raise HTMLConversionError(f"HTML内容生成失败: {e}", context=context, original_error=e)
    
    def _generate_html_header(self, theme_config: Dict, title: str, include_styles: bool) -> str:
        """生成HTML头部"""
        header_parts = [
            '<!DOCTYPE html>',
            '<html lang="zh-CN">',
            '<head>',
            '<meta charset="UTF-8">',
            '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
            f'<title>{title}</title>'
        ]
        
        if include_styles:
            # 基础样式
            header_parts.append('<style>')
            header_parts.append(f"body {{ {theme_config['body_style']} }}")
            header_parts.append(f"table {{ {theme_config['table_style']} }}")
            header_parts.append(f"td, th {{ {theme_config['cell_style']} }}")
            header_parts.append(f"th {{ {theme_config['header_style']} }}")
            
            # 附加样式
            additional_styles = self._get_additional_styles()
            header_parts.extend(additional_styles)
            
            header_parts.append('</style>')
        
        header_parts.append('</head>')
        return '\n'.join(header_parts)
    
    def _get_additional_styles(self) -> List[str]:
        """获取附加CSS样式"""
        return [
            # 合并单元格样式
            '.merged-cell { background-color: #f0f8ff; }',
            
            # 注释样式
            '.comment-cell { position: relative; cursor: help; }',
            '.comment-cell::after { content: "📝"; position: absolute; top: 2px; right: 2px; font-size: 10px; color: #666; }',
            '.comment-tooltip { display: none; position: absolute; top: 100%; left: 0; background: #333; color: #fff; padding: 8px 12px; border-radius: 4px; font-size: 12px; z-index: 1000; white-space: nowrap; box-shadow: 0 2px 8px rgba(0,0,0,0.3); }',
            '.comment-tooltip::before { content: ""; position: absolute; top: -5px; left: 10px; border-left: 5px solid transparent; border-right: 5px solid transparent; border-bottom: 5px solid #333; }',
            '.comment-cell:hover .comment-tooltip { display: block; }',
            
            # 超链接样式
            '.hyperlink-cell a { color: #0066cc; text-decoration: underline; }',
            '.hyperlink-cell a:hover { color: #0052a3; }',
            
            # 公式样式
            '.formula-cell { position: relative; }',
            '.formula-indicator { font-size: 10px; color: #666; position: absolute; top: 1px; left: 2px; }',
            '.formula-result { font-weight: normal; }',
            '.formula-error { background-color: #ffe6e6; color: #cc0000; }',
            '.formula-tooltip { display: none; position: absolute; background: #333; color: #fff; padding: 3px 6px; border-radius: 3px; font-size: 11px; z-index: 1000; }',
            '.formula-cell:hover .formula-tooltip { display: block; }',
            
            # 数据类型样式
            '.number-cell { text-align: right; }',
            '.date-cell { text-align: center; color: #666; }',
            '.text-cell { text-align: left; }',
            
            # 响应式设计
            '@media (max-width: 768px) {',
            '  table { font-size: 12px; }',
            '  td, th { padding: 4px; }',
            '}',
            
            # 打印样式
            '@media print {',
            '  body { margin: 0; background: white; }',
            '  table { page-break-inside: avoid; }',
            '}'
        ]
    
    @error_handler(operation="工作表HTML转换")
    def _convert_sheet_to_html(self, sheet: Dict, theme_config: Dict, include_styles: bool) -> str:
        """将单个工作表转换为HTML表格"""
        sheet_name = sheet.get('sheet_name', 'Sheet')
        data = sheet.get('data', [])
        styles = sheet.get('styles', [])
        merged_cells = sheet.get('merged_cells', [])
        comments = sheet.get('comments', {})
        hyperlinks = sheet.get('hyperlinks', {})
        
        if not data:
            return f'<h2>{sheet_name}</h2><p>表格为空</p>'
        
        html_parts = []
        html_parts.append(f'<h2>{sheet_name}</h2>')
        html_parts.append('<table>')
        
        # 创建合并单元格映射
        merged_map = self._create_merged_map(merged_cells)
        
        # 处理每一行
        for row_idx, row_data in enumerate(data):
            if not any(cell.strip() for cell in row_data if isinstance(cell, str)):
                continue  # 跳过空行
            
            html_parts.append('<tr>')
            
            for col_idx, cell_value in enumerate(row_data):
                # 检查是否是被合并的单元格（不是起始单元格）
                if (row_idx, col_idx) in merged_map and merged_map[(row_idx, col_idx)] != (row_idx, col_idx):
                    continue
                
                # 获取单元格样式和属性
                cell_html = self._create_cell_html(
                    cell_value, row_idx, col_idx, styles, comments, hyperlinks, 
                    merged_map, include_styles
                )
                html_parts.append(cell_html)
            
            html_parts.append('</tr>')
        
        html_parts.append('</table>')
        return '\n'.join(html_parts)
    
    def _create_merged_map(self, merged_cells: List[Tuple[int, int, int, int]]) -> Dict:
        """创建合并单元格映射"""
        merged_map = {}
        for start_row, start_col, end_row, end_col in merged_cells:
            for r in range(start_row, end_row + 1):
                for c in range(start_col, end_col + 1):
                    merged_map[(r, c)] = (start_row, start_col)
        return merged_map
    
    def _create_cell_html(self, cell_value: str, row_idx: int, col_idx: int,
                         styles: List[List[Dict]], comments: Dict, hyperlinks: Dict,
                         merged_map: Dict, include_styles: bool) -> str:
        """创建单个单元格的HTML"""
        cell_key = f"{row_idx}_{col_idx}"
        style_info = {}
        
        # 获取样式信息
        if (include_styles and row_idx < len(styles) and 
            col_idx < len(styles[row_idx])):
            style_info = styles[row_idx][col_idx] or {}
        
        # 构建单元格属性
        cell_attrs = []
        css_classes = []
        css_styles = []
        
        # 处理合并单元格
        if (row_idx, col_idx) in merged_map:
            start_row, start_col = merged_map[(row_idx, col_idx)]
            if start_row == row_idx and start_col == col_idx:
                # 这是合并单元格的起始单元格
                rowspan, colspan = self._calculate_span(merged_map, row_idx, col_idx)
                if rowspan > 1:
                    cell_attrs.append(f'rowspan="{rowspan}"')
                if colspan > 1:
                    cell_attrs.append(f'colspan="{colspan}"')
                css_classes.append('merged-cell')
        
        # 应用样式
        if include_styles and style_info:
            cell_styles = self._apply_cell_styles(style_info)
            css_styles.extend(cell_styles)
        
        # 处理公式
        cell_content, formula_tooltip = self._process_formula_content(cell_value, style_info)
        is_formula_html = '<span' in cell_content  # 检查是否包含HTML标签
        
        # 处理超链接
        if cell_key in hyperlinks:
            css_classes.append('hyperlink-cell')
            url = hyperlinks[cell_key]
            if is_formula_html:
                # 公式内容已经是HTML，不需要转义
                cell_content = f'<a href="{html.escape(url)}" target="_blank">{cell_content}</a>'
            else:
                cell_content = f'<a href="{html.escape(url)}" target="_blank">{html.escape(cell_content)}</a>'
        else:
            # 对于公式内容，如果已经包含HTML标签，则不转义
            if not is_formula_html:
                cell_content = html.escape(cell_content)
        
        # 处理注释
        comment_html = ''
        if cell_key in comments and self.config.INCLUDE_COMMENTS:
            css_classes.append('comment-cell')
            comment_html = f'<div class="comment-tooltip">{html.escape(comments[cell_key])}</div>'
        
        # 组装单元格
        tag = 'th' if row_idx == 0 else 'td'
        
        # 构建class属性
        if css_classes:
            class_str = ' '.join(css_classes)
            cell_attrs.append(f'class="{class_str}"')
        # 构建style属性
        if css_styles:
            style_str = '; '.join(css_styles)
            cell_attrs.append(f'style="{style_str}"')
        
        # 只添加公式tooltip到title属性（评论使用CSS tooltip）
        if formula_tooltip:
            cell_attrs.append(f'title="{html.escape(formula_tooltip)}"')
        
        # 构建属性字符串
        if cell_attrs:
            attrs_str = ' ' + ' '.join(cell_attrs)
        else:
            attrs_str = ''
        return f'<{tag}{attrs_str}>{cell_content}{comment_html}</{tag}>'
    
    def _calculate_span(self, merged_map: Dict, row_idx: int, col_idx: int) -> Tuple[int, int]:
        """计算合并单元格的跨度"""
        max_row = row_idx
        max_col = col_idx
        
        # 查找最大行和列
        for (r, c), (start_r, start_c) in merged_map.items():
            if start_r == row_idx and start_c == col_idx:
                max_row = max(max_row, r)
                max_col = max(max_col, c)
        
        return max_row - row_idx + 1, max_col - col_idx + 1
    
    def _apply_cell_styles(self, style_info: Dict) -> List[str]:
        """应用单元格样式"""
        styles = []
        
        # 字体样式
        if style_info.get('bold'):
            styles.append('font-weight: bold')
        if style_info.get('italic'):
            styles.append('font-style: italic')
        if style_info.get('font_size'):
            styles.append(f'font-size: {style_info["font_size"]}px')
        if style_info.get('font_name'):
            styles.append(f'font-family: "{style_info["font_name"]}", sans-serif')
        if style_info.get('font_color'):
            styles.append(f'color: {style_info["font_color"]}')
        if style_info.get('underline'):
            styles.append('text-decoration: underline')
        if style_info.get('strike'):
            styles.append('text-decoration: line-through')
        
        # 背景色
        if style_info.get('bg_color'):
            styles.append(f'background-color: {style_info["bg_color"]}')
        
        # 对齐
        if style_info.get('align'):
            align_map = self.config.ALIGNMENT_MAPPING
            align_value = align_map.get(style_info['align'], style_info['align'])
            styles.append(f'text-align: {align_value}')
        
        if style_info.get('valign'):
            valign_map = self.config.VERTICAL_ALIGNMENT_MAPPING
            valign_value = valign_map.get(style_info['valign'], style_info['valign'])
            styles.append(f'vertical-align: {valign_value}')
        
        # 文本换行
        if style_info.get('wrap_text'):
            styles.append('white-space: pre-wrap')
        
        # 边框
        if style_info.get('border'):
            border_styles = self._convert_border_styles(style_info['border'])
            styles.extend(border_styles)
        
        return styles
    
    def _convert_border_styles(self, border_info: Dict) -> List[str]:
        """转换边框样式"""
        styles = []
        style_map = self.config.BORDER_STYLE_MAPPING
        
        for side, border_detail in border_info.items():
            if isinstance(border_detail, dict):
                border_style = border_detail.get('style', 'solid')
                border_color = border_detail.get('color', '#000000')
                
                # 映射边框样式
                css_style = style_map.get(border_style, '1px solid')
                styles.append(f'border-{side}: {css_style} {border_color}')
        
        return styles
    
    def _process_formula_content(self, cell_value: str, style_info: Dict) -> Tuple[str, str]:
        """处理公式内容"""
        # 导入FormulaInfo以避免循环导入
        from .formula_processor import FormulaInfo, FormulaError
        
        # 检查是否包含公式信息
        if isinstance(cell_value, FormulaInfo):
            formula_info = cell_value
            
            # 构建显示内容
            classes = ['formula-cell']
            if formula_info.error:
                classes.append('formula-error')
            
            # 构建tooltip
            formula_type = getattr(formula_info, 'formula_type', "未知")
            if hasattr(formula_type, 'value'):
                formula_type = formula_type.value
            
            tooltip_parts = [f"公式: {formula_info.original_formula} ({formula_type})"]
            
            if formula_info.error:
                error_value = formula_info.error.value if hasattr(formula_info.error, 'value') else str(formula_info.error)
                tooltip_parts.append(f"[错误: {error_value}]")
            
            tooltip = " ".join(tooltip_parts)
            
            # 构建显示内容
            parts = ['<span class="formula-indicator">ƒ</span>']
            
            if formula_info.calculated_value is not None:
                # 格式化计算结果
                result = formula_info.calculated_value
                if isinstance(result, float):
                    if result.is_integer():
                        result_str = str(int(result))
                    else:
                        result_str = f"{result:.6g}"
                else:
                    result_str = str(result)
                parts.append(f'<span class="formula-result">{result_str}</span>')
            elif formula_info.error:
                error_value = formula_info.error.value if hasattr(formula_info.error, 'value') else str(formula_info.error)
                parts.append(f'<span class="formula-result">{error_value}</span>')
            else:
                parts.append('<span class="formula-result">#ERROR</span>')
            
            content = ''.join(parts)
            return content, tooltip
        
        # 检查是否是公式字符串
        if isinstance(cell_value, str) and cell_value.startswith('='):
            tooltip = f"公式: {cell_value}"
            content = f'<span class="formula-indicator">ƒ</span><span class="formula-result">{cell_value[1:]}</span>'
            return content, tooltip
        
        # 普通单元格
        return str(cell_value) if cell_value is not None else '', ''
    
    def _write_html_file(self, output_path: str, html_content: str, context: ErrorContext):
        """写入HTML文件"""
        try:
            encoding = getattr(self.config, 'HTML_DEFAULT_ENCODING', 'utf-8')
            with open(output_path, 'w', encoding=encoding) as f:
                f.write(html_content)
            
            self.logger.info(f"HTML文件写入成功: {output_path}")
            
        except PermissionError as e:
            raise SecurityError(f"没有权限写入文件 {output_path}: {e}", context=context, original_error=e)
        except Exception as e:
            raise HTMLConversionError(f"写入HTML文件失败 {output_path}: {e}", context=context, original_error=e)
    
    def get_error_summary(self):
        """获取转换过程中的错误统计"""
        return self.error_handler.get_error_summary()
