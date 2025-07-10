# html_converter.py
# HTML转换与样式映射

class HTMLConverter:
    def __init__(self, table_data):
        """
        table_data: SheetParser 解析得到的单个sheet结构化数据
        """
        self.table_data = table_data

    def _style_dict_to_str(self, style):
        css = []
        if style.get('bold'):
            css.append('font-weight:bold;')
        if style.get('italic'):
            css.append('font-style:italic;')
        if style.get('font_size'):
            css.append(f'font-size:{style["font_size"]}pt;')
        if style.get('font_color'):
            css.append(f'color:{style["font_color"]};')
        if style.get('bg_color'):
            css.append(f'background-color:{style["bg_color"]};')
        if style.get('align'):
            css.append(f'text-align:{style["align"]};')
        if style.get('valign'):
            css.append(f'vertical-align:{style["valign"]};')
        if style.get('font_name'):
            css.append(f'font-family:{style["font_name"]};')
        # 新增样式
        if style.get('underline'):
            css.append('text-decoration:underline;')
        if style.get('strike'):
            css.append('text-decoration:line-through;')
        if style.get('wrap_text'):
            css.append('white-space:pre-wrap;')
        # 边框样式
        if style.get('border'):
            border = style['border']
            for side, props in border.items():
                border_style = props['style']
                border_color = props['color']
                # 将Excel边框样式映射为CSS
                if border_style == 'thin':
                    css.append(f'border-{side}:1px solid {border_color};')
                elif border_style == 'medium':
                    css.append(f'border-{side}:2px solid {border_color};')
                elif border_style == 'thick':
                    css.append(f'border-{side}:3px solid {border_color};')
                elif border_style == 'dashed':
                    css.append(f'border-{side}:1px dashed {border_color};')
                elif border_style == 'dotted':
                    css.append(f'border-{side}:1px dotted {border_color};')
                else:
                    css.append(f'border-{side}:1px solid {border_color};')
        return ''.join(css)

    def to_html(self):
        """
        将结构化表格数据转换为HTML字符串，支持合并单元格和样式映射。
        """
        rows = self.table_data['rows']
        cols = self.table_data['cols']
        data = self.table_data['data']
        styles = self.table_data.get('styles', [])
        merged_cells = self.table_data.get('merged_cells', [])

        # 构建合并单元格映射
        merge_map = {}
        for (r1, c1, r2, c2) in merged_cells:
            for r in range(r1, r2+1):
                for c in range(c1, c2+1):
                    merge_map[(r, c)] = (r1, c1, r2, c2)

        # 生成CSS样式
        css = """
        <style>
        .comment-cell {
            position: relative;
        }
        .comment-indicator {
            position: absolute;
            top: 2px;
            right: 2px;
            width: 6px;
            height: 6px;
            background-color: #FF0000;
            border-radius: 50%;
            z-index: 10;
        }
        .comment-tooltip {
            position: absolute;
            background-color: #FFFFD1;
            border: 1px solid #000000;
            padding: 8px;
            border-radius: 4px;
            font-size: 12px;
            max-width: 200px;
            white-space: pre-wrap;
            z-index: 1000;
            display: none;
            box-shadow: 2px 2px 5px rgba(0,0,0,0.3);
        }
        .comment-tooltip::before {
            content: '';
            position: absolute;
            top: -5px;
            left: 10px;
            width: 0;
            height: 0;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-bottom: 5px solid #FFFFD1;
        }
        </style>
        """

        html = [css, '<table border="1" cellspacing="0" cellpadding="4">']
        comment_counter = 0
        
        for r in range(rows):
            html.append('<tr>')
            for c in range(cols):
                # 跳过被合并的单元格（只保留左上角）
                if (r, c) in merge_map and (r, c) != (merge_map[(r, c)][0], merge_map[(r, c)][1]):
                    continue
                cell = data[r][c] if r < len(data) and c < len(data[r]) else ''
                style = styles[r][c] if r < len(styles) and c < len(styles[r]) else {}
                style_str = self._style_dict_to_str(style)
                attrs = ''
                if style_str:
                    attrs += f' style="{style_str}"'
                
                # 批注处理
                comment_html = ''
                if style.get('comment'):
                    comment_counter += 1
                    comment_id = f'comment_{comment_counter}'
                    comment_text = style['comment'].replace('"', '&quot;').replace("'", "&#39;")
                    attrs += f' class="comment-cell" onmouseover="showComment(\'{comment_id}\')" onmouseout="hideComment(\'{comment_id}\')"'
                    comment_html = f'<div class="comment-indicator"></div><div id="{comment_id}" class="comment-tooltip">{comment_text}</div>'
                
                # 超链接处理
                if style.get('hyperlink'):
                    cell_html = f'<a href="{style["hyperlink"]}" target="_blank">{cell}</a>'
                else:
                    cell_html = cell
                
                # 合并单元格处理
                if (r, c) in merge_map and (r, c) == (merge_map[(r, c)][0], merge_map[(r, c)][1]):
                    r1, c1, r2, c2 = merge_map[(r, c)]
                    rowspan = r2 - r1 + 1
                    colspan = c2 - c1 + 1
                    if rowspan > 1:
                        attrs += f' rowspan="{rowspan}"'
                    if colspan > 1:
                        attrs += f' colspan="{colspan}"'
                
                # 第一行为表头
                if r == 0:
                    html.append(f'<th{attrs}>{cell_html}{comment_html}</th>')
                else:
                    html.append(f'<td{attrs}>{cell_html}{comment_html}</td>')
            html.append('</tr>')
        
        html.append('</table>')
        
        # 添加JavaScript
        js = """
        <script>
        function showComment(commentId) {
            var tooltip = document.getElementById(commentId);
            if (tooltip) {
                tooltip.style.display = 'block';
            }
        }
        function hideComment(commentId) {
            var tooltip = document.getElementById(commentId);
            if (tooltip) {
                tooltip.style.display = 'none';
            }
        }
        </script>
        """
        
        return '\n'.join(html) + js
