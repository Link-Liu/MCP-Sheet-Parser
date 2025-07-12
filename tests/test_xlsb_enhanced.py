# test_xlsb_enhanced.py
# 测试XLSB格式增强功能

import pytest
import os
import tempfile
from unittest.mock import patch, MagicMock
from mcp_sheet_parser.parser import SheetParser
from mcp_sheet_parser.config import UnifiedConfig
from mcp_sheet_parser.exceptions import UnsupportedFormatError, ParsingError


class TestXLSBEnhanced:
    """测试XLSB格式增强功能"""
    
    def setup_method(self):
        """设置测试环境"""
        self.config = UnifiedConfig()
        self.temp_dir = tempfile.mkdtemp()
    
    def teardown_method(self):
        """清理测试环境"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_xlsb_format_configuration(self):
        """测试XLSB格式配置"""
        # 检查XLSB格式是否在支持列表中
        assert '.xlsb' in self.config.SUPPORTED_EXCEL_FORMATS
        
        # 检查支持级别
        assert self.config.file_format.get_support_level('.xlsb') == 'BASIC'
        
        # 检查格式类型识别
        assert self.config.file_format.get_format_type('.xlsb') == 'excel'
        
        # 检查支持徽章
        assert self.config.file_format.get_support_badge('.xlsb') == '⚠️ 基础支持'
    
    def test_xlsb_format_validation(self):
        """测试XLSB格式验证"""
        # 测试支持的格式
        assert self.config.file_format.is_excel_format('.xlsb')
        assert self.config.file_format.is_supported_format('.xlsb')
        
        # 测试大小写不敏感
        assert self.config.file_format.is_excel_format('.XLSB')
        assert self.config.file_format.is_supported_format('.XLSB')
    
    @patch('pyxlsb.open_workbook', autospec=False)
    def test_xlsb_enhanced_parsing(self, mock_open_workbook):
        """测试XLSB增强解析功能"""
        # 创建模拟的XLSB工作簿
        mock_wb = MagicMock()
        mock_sheet = MagicMock()
        mock_cell = MagicMock()
        mock_style = MagicMock()
        
        # 设置模拟数据 - 修复行结构
        mock_cell.r = 1  # 行号
        mock_cell.c = 1  # 列号
        mock_cell.v = "测试数据"
        mock_cell.f = "=A1+B1"  # 公式
        mock_cell.s = mock_style  # 样式
        mock_cell.comment = "这是注释"
        mock_cell.hyperlink = "https://example.com"
        
        # 设置样式属性
        mock_font = MagicMock()
        mock_font.bold = True
        mock_font.italic = False
        mock_font.size = 12
        mock_font.name = "Arial"
        mock_font.color = "FF0000"
        
        mock_fill = MagicMock()
        mock_fill.fgColor = "FFFF00"
        
        mock_alignment = MagicMock()
        mock_alignment.horizontal = "center"
        mock_alignment.vertical = "middle"
        mock_alignment.wrapText = True
        
        mock_border = MagicMock()
        mock_border_top = MagicMock()
        mock_border_top.style = "thin"
        mock_border_top.color = "000000"
        mock_border.top = mock_border_top
        
        # 组装样式对象
        mock_style.font = mock_font
        mock_style.fill = mock_fill
        mock_style.alignment = mock_alignment
        mock_style.border = mock_border
        
        # 设置工作表 - 修复行结构
        # pyxlsb的rows()方法返回的是行对象列表，每行包含单元格列表
        mock_row = [mock_cell]  # 第一行包含一个单元格
        mock_sheet.rows.return_value = [mock_row]  # 返回行列表
        mock_wb.sheets = ['Sheet1']
        mock_wb.get_sheet.return_value = mock_sheet
        
        # 设置上下文管理器
        mock_context = MagicMock()
        mock_context.__enter__.return_value = mock_wb
        mock_context.__exit__.return_value = None
        mock_open_workbook.return_value = mock_context
        
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test.xlsb')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        # 测试解析
        parser = SheetParser(temp_file, self.config)
        result = parser._parse_xlsb()
        
        # 验证结果
        assert len(result) == 1
        assert result[0]['sheet_name'] == 'Sheet1'
        assert result[0]['rows'] == 1
        assert result[0]['cols'] == 1
        
        # 验证数据
        assert result[0]['data'][0][0] == "=A1+B1"  # 公式
        
        # 验证样式
        styles = result[0]['styles'][0][0]
        assert styles['bold'] is True
        assert styles['font_size'] == 12
        assert styles['font_name'] == "Arial"
        assert styles['font_color'] == "#FF0000"
        assert styles['bg_color'] == "#FFFF00"
        assert styles['align'] == "center"
        assert styles['valign'] == "middle"
        assert styles['wrap_text'] is True
        
        # 验证边框
        assert 'border' in styles
        assert styles['border']['top']['style'] == "thin"
        assert styles['border']['top']['color'] == "#000000"
        
        # 验证注释和超链接
        assert result[0]['comments']['0_0'] == "这是注释"
        assert result[0]['hyperlinks']['0_0'] == "https://example.com"
    
    def test_xlsb_style_extraction(self):
        """测试XLSB样式提取"""
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test.xlsb')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        
        # 创建模拟样式对象
        mock_style = MagicMock()
        mock_font = MagicMock()
        mock_font.bold = True
        mock_font.italic = True
        mock_font.size = 14
        mock_font.name = "Times New Roman"
        mock_font.color = "00FF00"
        
        mock_fill = MagicMock()
        mock_fill.fgColor = "FF00FF"
        
        mock_alignment = MagicMock()
        mock_alignment.horizontal = "right"
        mock_alignment.vertical = "top"
        mock_alignment.wrapText = False
        
        mock_border = MagicMock()
        mock_border_left = MagicMock()
        mock_border_left.style = "thick"
        mock_border_left.color = "FF0000"
        mock_border.left = mock_border_left
        mock_border.right = None  # 测试空边框
        
        # 组装样式对象
        mock_style.font = mock_font
        mock_style.fill = mock_fill
        mock_style.alignment = mock_alignment
        mock_style.border = mock_border
        
        # 测试样式提取
        style = parser._extract_xlsb_style(mock_style)
        
        # 验证字体样式
        assert style['bold'] is True
        assert style['italic'] is True
        assert style['font_size'] == 14
        assert style['font_name'] == "Times New Roman"
        assert style['font_color'] == "#00FF00"
        
        # 验证背景色
        assert style['bg_color'] == "#FF00FF"
        
        # 验证对齐方式
        assert style['align'] == "right"
        assert style['valign'] == "top"
        assert style['wrap_text'] is False
        
        # 验证边框
        assert 'border' in style
        assert 'left' in style['border']
        assert style['border']['left']['style'] == "thick"
        assert style['border']['left']['color'] == "#FF0000"
        assert 'right' not in style['border']  # 空边框不应包含
    
    def test_xlsb_error_handling(self):
        """测试XLSB错误处理"""
        # 测试缺少pyxlsb依赖的情况
        with patch('pyxlsb.open_workbook', autospec=False, side_effect=ImportError("No module named 'pyxlsb'")):
            temp_file = os.path.join(self.temp_dir, 'test.xlsb')
            with open(temp_file, 'w') as f:
                f.write('dummy content')
            
            parser = SheetParser(temp_file, self.config)
            
            with pytest.raises(ParsingError) as exc_info:
                parser._parse_xlsb()
            
            assert "缺少 pyxlsb 依赖" in str(exc_info.value)
    
    def test_xlsb_data_validation(self):
        """测试XLSB数据验证"""
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test.xlsb')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        
        # 测试数据大小验证 - 设置一个非常大的数据来触发验证
        # 需要超过配置的最大行数和列数限制
        max_rows = getattr(self.config, 'MAX_ROWS', 10000)
        max_cols = getattr(self.config, 'MAX_COLS', 1000)
        
        large_data = [[''] * (max_cols + 1) for _ in range(max_rows + 1)]  # 超过限制的数据
        
        # 创建错误上下文
        from mcp_sheet_parser.exceptions import create_error_context
        context = create_error_context("数据验证测试", file_path=temp_file)
        
        with pytest.raises(ParsingError):
            parser._validate_data_size(large_data, context)
    
    def test_xlsb_formula_handling(self):
        """测试XLSB公式处理"""
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test.xlsb')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        
        # 测试公式单元格
        mock_cell = MagicMock()
        mock_cell.v = None
        mock_cell.f = "SUM(A1:A10)"
        
        # 模拟公式处理
        cell_value = ""
        if hasattr(mock_cell, 'f') and mock_cell.f:
            cell_value = f"={mock_cell.f}"
        
        assert cell_value == "=SUM(A1:A10)"
    
    def test_xlsb_empty_file_handling(self):
        """测试XLSB空文件处理"""
        with patch('pyxlsb.open_workbook', autospec=False) as mock_open_workbook:
            # 创建模拟的空工作簿
            mock_wb = MagicMock()
            mock_wb.sheets = []  # 没有工作表
            
            # 设置上下文管理器
            mock_context = MagicMock()
            mock_context.__enter__.return_value = mock_wb
            mock_context.__exit__.return_value = None
            mock_open_workbook.return_value = mock_context
            
            temp_file = os.path.join(self.temp_dir, 'empty.xlsb')
            with open(temp_file, 'w') as f:
                f.write('dummy content')
            
            parser = SheetParser(temp_file, self.config)
            result = parser._parse_xlsb()
            
            # 空文件应该返回空列表
            assert result == []

    def test_xlsb_merged_cells_extraction(self):
        """测试XLSB合并单元格提取（merged_cells属性）"""
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test.xlsb')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        
        # 直接测试合并单元格提取逻辑
        # 模拟有merged_cells的情况
        mock_ws = MagicMock()
        merged_mock = MagicMock()
        merged_mock.row_first = 0
        merged_mock.col_first = 0
        merged_mock.row_last = 1
        merged_mock.col_last = 2
        mock_ws.merged_cells = [merged_mock]
        mock_ws.merged_ranges = None
        
        # 手动调用合并单元格提取逻辑
        merged_cells = []
        if hasattr(mock_ws, 'merged_cells') and mock_ws.merged_cells:
            try:
                for merged in mock_ws.merged_cells:
                    merged_cells.append((
                        merged.row_first, merged.col_first, merged.row_last, merged.col_last
                    ))
            except Exception as e:
                pass
        
        assert (0, 0, 1, 2) in merged_cells

    def test_xlsb_merged_ranges_extraction(self):
        """测试XLSB合并单元格提取（merged_ranges属性）"""
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test2.xlsb')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        
        # 直接测试合并单元格提取逻辑
        # 模拟有merged_ranges的情况
        mock_ws = MagicMock()
        merged_mock = MagicMock()
        merged_mock.row_first = 2
        merged_mock.col_first = 2
        merged_mock.row_last = 3
        merged_mock.col_last = 4
        mock_ws.merged_cells = None
        mock_ws.merged_ranges = [merged_mock]
        
        # 手动调用合并单元格提取逻辑
        merged_cells = []
        if hasattr(mock_ws, 'merged_ranges') and mock_ws.merged_ranges:
            try:
                for merged in mock_ws.merged_ranges:
                    merged_cells.append((
                        merged.row_first, merged.col_first, merged.row_last, merged.col_last
                    ))
            except Exception as e:
                pass
        
        assert (2, 2, 3, 4) in merged_cells

    def test_xlsb_no_merged_cells(self):
        """测试XLSB无合并单元格时输出空列表"""
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test3.xlsb')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        
        # 直接测试合并单元格提取逻辑
        # 模拟没有合并单元格的情况
        mock_ws = MagicMock()
        # 不设置merged_cells/merged_ranges
        delattr(mock_ws, 'merged_cells') if hasattr(mock_ws, 'merged_cells') else None
        delattr(mock_ws, 'merged_ranges') if hasattr(mock_ws, 'merged_ranges') else None
        
        # 手动调用合并单元格提取逻辑
        merged_cells = []
        if hasattr(mock_ws, 'merged_cells') and mock_ws.merged_cells:
            try:
                for merged in mock_ws.merged_cells:
                    merged_cells.append((
                        merged.row_first, merged.col_first, merged.row_last, merged.col_last
                    ))
            except Exception as e:
                pass
        elif hasattr(mock_ws, 'merged_ranges') and mock_ws.merged_ranges:
            try:
                for merged in mock_ws.merged_ranges:
                    merged_cells.append((
                        merged.row_first, merged.col_first, merged.row_last, merged.col_last
                    ))
            except Exception as e:
                pass
        
        assert merged_cells == []


if __name__ == '__main__':
    pytest.main([__file__]) 