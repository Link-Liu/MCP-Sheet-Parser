# parser.py
# 表格解析核心逻辑

import os
import pandas as pd
import openpyxl
from pandas.errors import EmptyDataError
from openpyxl.styles.colors import Color

def _get_cell_style(cell):
    style = {}
    # 字体
    if cell.font:
        style['bold'] = cell.font.bold
        style['italic'] = cell.font.italic
        style['font_size'] = cell.font.sz
        style['font_name'] = cell.font.name
        style['underline'] = cell.font.underline
        style['strike'] = cell.font.strike
        # 字体颜色
        if cell.font.color and cell.font.color.type == 'rgb' and cell.font.color.rgb:
            style['font_color'] = f"#{cell.font.color.rgb[-6:]}"
    # 背景色
    if cell.fill and hasattr(cell.fill, 'fgColor') and cell.fill.fgColor.type == 'rgb' and cell.fill.fgColor.rgb and cell.fill.fgColor.rgb != '00000000':
        style['bg_color'] = f"#{cell.fill.fgColor.rgb[-6:]}"
    # 对齐
    if cell.alignment:
        if cell.alignment.horizontal:
            style['align'] = cell.alignment.horizontal
        if cell.alignment.vertical:
            style['valign'] = cell.alignment.vertical
        style['wrap_text'] = cell.alignment.wrap_text
    # 边框
    if cell.border:
        border = {}
        for side in ['top', 'bottom', 'left', 'right']:
            border_side = getattr(cell.border, side)
            if border_side and border_side.style and border_side.style != 'none':
                border[side] = {
                    'style': border_side.style,
                    'color': f"#{border_side.color.rgb[-6:]}" if border_side.color and border_side.color.rgb else '#000000'
                }
        if border:
            style['border'] = border
    # 超链接
    if cell.hyperlink:
        style['hyperlink'] = str(cell.hyperlink.target)
    # 批注/注释
    if cell.comment:
        style['comment'] = cell.comment.text
    return style

class SheetParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.sheets = []  # 存储所有sheet的结构化数据

    def parse(self):
        """
        解析表格文件，返回结构化数据：
        [
            {
                'sheet_name': str,
                'rows': int,
                'cols': int,
                'data': [[cell, ...], ...],
                'styles': [[style_dict, ...], ...],
                'merged_cells': [(start_row, start_col, end_row, end_col), ...]
            },
            ...
        ]
        """
        ext = os.path.splitext(self.file_path)[-1].lower()
        if ext == '.csv':
            return [self._parse_csv()]
        elif ext in ['.xlsx', '.xls']:
            return self._parse_excel()
        else:
            raise ValueError(f'暂不支持的文件格式: {ext}')

    def _parse_csv(self):
        try:
            df = pd.read_csv(self.file_path, header=None, dtype=str)
            data = df.fillna('').values.tolist()
        except EmptyDataError:
            data = []
        # CSV无样式
        styles = [[{} for _ in row] for row in data]
        return {
            'sheet_name': os.path.basename(self.file_path),
            'rows': len(data),
            'cols': len(data[0]) if data else 0,
            'data': data,
            'styles': styles,
            'merged_cells': []  # CSV无合并单元格
        }

    def _parse_excel(self):
        wb = openpyxl.load_workbook(self.file_path, data_only=True)
        sheets = []
        for ws in wb.worksheets:
            data = []
            styles = []
            # 使用max_row和max_column确保遍历所有行和列
            for r in range(1, ws.max_row + 1):
                data_row = []
                style_row = []
                for c in range(1, ws.max_column + 1):
                    cell = ws.cell(row=r, column=c)
                    data_row.append(cell.value if cell.value is not None else '')
                    style_row.append(_get_cell_style(cell))
                data.append(data_row)
                styles.append(style_row)
            merged_cells = []
            for merged in ws.merged_cells.ranges:
                merged_cells.append((merged.min_row-1, merged.min_col-1, merged.max_row-1, merged.max_col-1))
            sheets.append({
                'sheet_name': ws.title,
                'rows': ws.max_row,
                'cols': ws.max_column,
                'data': data,
                'styles': styles,
                'merged_cells': merged_cells
            })
        return sheets
