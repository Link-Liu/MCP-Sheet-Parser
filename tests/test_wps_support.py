# test_wps_support.py
# 测试WPS格式支持功能

import pytest
import os
import tempfile
from unittest.mock import patch, MagicMock
from mcp_sheet_parser.parser import SheetParser
from mcp_sheet_parser.config import UnifiedConfig
from mcp_sheet_parser.exceptions import UnsupportedFormatError, ParsingError


class TestWPSSupport:
    """测试WPS格式支持"""
    
    def setup_method(self):
        """设置测试环境"""
        self.config = UnifiedConfig()
        self.temp_dir = tempfile.mkdtemp()
    
    def teardown_method(self):
        """清理测试环境"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_wps_format_configuration(self):
        """测试WPS格式配置"""
        # 检查WPS格式是否在支持列表中
        assert '.et' in self.config.SUPPORTED_WPS_FORMATS
        assert '.ett' in self.config.SUPPORTED_WPS_FORMATS
        assert '.ets' in self.config.SUPPORTED_WPS_FORMATS
        
        # 检查支持级别
        assert self.config.file_format.get_support_level('.et') == 'BASIC'
        assert self.config.file_format.get_support_level('.ett') == 'BASIC'
        assert self.config.file_format.get_support_level('.ets') == 'BASIC'
        
        # 检查格式类型识别
        assert self.config.file_format.get_format_type('.et') == 'wps'
        assert self.config.file_format.get_format_type('.ett') == 'wps'
        assert self.config.file_format.get_format_type('.ets') == 'wps'
        
        # 检查支持徽章
        assert self.config.file_format.get_support_badge('.et') == '⚠️ 基础支持'
        assert self.config.file_format.get_support_badge('.ett') == '⚠️ 基础支持'
        assert self.config.file_format.get_support_badge('.ets') == '⚠️ 基础支持'
    
    def test_wps_format_validation(self):
        """测试WPS格式验证"""
        # 测试支持的格式
        assert self.config.file_format.is_wps_format('.et')
        assert self.config.file_format.is_wps_format('.ett')
        assert self.config.file_format.is_wps_format('.ets')
        
        # 测试不支持的格式
        assert not self.config.file_format.is_wps_format('.xlsx')
        assert not self.config.file_format.is_wps_format('.csv')
        
        # 测试大小写不敏感
        assert self.config.file_format.is_wps_format('.ET')
        assert self.config.file_format.is_wps_format('.ETT')
        assert self.config.file_format.is_wps_format('.ETS')
    
    @patch('mcp_sheet_parser.parser.openpyxl.load_workbook')
    def test_wps_workbook_parsing(self, mock_load_workbook):
        """测试WPS工作簿解析"""
        # 创建模拟的Excel工作簿
        mock_wb = MagicMock()
        mock_ws = MagicMock()
        mock_cell = MagicMock()
        
        mock_ws.title = 'Sheet1'  # 修复sheet_name断言
        # 设置模拟数据
        mock_ws.max_row = 3
        mock_ws.max_column = 3
        mock_ws.cell.return_value = mock_cell
        mock_cell.value = "测试数据"
        mock_cell.data_type = 's'
        
        mock_wb.worksheets = [mock_ws]
        mock_wb.sheetnames = ['Sheet1']
        mock_load_workbook.return_value = mock_wb
        
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test.et')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        # 测试解析
        parser = SheetParser(temp_file, self.config)
        result = parser._parse_wps_workbook()
        
        # 验证结果
        assert len(result) == 1
        assert result[0]['sheet_name'] == 'Sheet1'
        assert result[0]['wps_format'] == 'workbook'
        assert 'wps_metadata' in result[0]
    
    @patch('mcp_sheet_parser.parser.openpyxl.load_workbook')
    def test_wps_template_parsing(self, mock_load_workbook):
        """测试WPS模板解析"""
        # 创建模拟的Excel工作簿
        mock_wb = MagicMock()
        mock_ws = MagicMock()
        mock_cell = MagicMock()
        
        mock_ws.title = 'Template'  # 修复sheet_name断言
        # 设置模拟数据，包含模板变量
        mock_ws.max_row = 3
        mock_ws.max_column = 3
        mock_ws.cell.return_value = mock_cell
        mock_cell.value = "{{变量名}}"
        mock_cell.data_type = 's'
        
        mock_wb.worksheets = [mock_ws]
        mock_wb.sheetnames = ['Template']
        mock_load_workbook.return_value = mock_wb
        
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test.ett')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        # 测试解析
        parser = SheetParser(temp_file, self.config)
        result = parser._parse_wps_template()
        
        # 验证结果
        assert len(result) == 1
        assert result[0]['sheet_name'] == 'Template'
        assert result[0]['wps_format'] == 'template'
        assert result[0]['is_template'] is True
        assert 'template_variables' in result[0]
    
    @patch('mcp_sheet_parser.parser.openpyxl.load_workbook')
    def test_wps_backup_parsing(self, mock_load_workbook):
        """测试WPS备份解析"""
        # 创建模拟的Excel工作簿
        mock_wb = MagicMock()
        mock_ws = MagicMock()
        mock_cell = MagicMock()
        
        mock_ws.title = 'Backup'  # 修复sheet_name断言
        # 设置模拟数据
        mock_ws.max_row = 3
        mock_ws.max_column = 3
        mock_ws.cell.return_value = mock_cell
        mock_cell.value = "备份数据"
        mock_cell.data_type = 's'
        
        mock_wb.worksheets = [mock_ws]
        mock_wb.sheetnames = ['Backup']
        mock_load_workbook.return_value = mock_wb
        
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test_backup.ets')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        # 测试解析
        parser = SheetParser(temp_file, self.config)
        result = parser._parse_wps_backup()
        
        # 验证结果
        assert len(result) == 1
        assert result[0]['sheet_name'] == 'Backup'
        assert result[0]['wps_format'] == 'backup'
        assert result[0]['is_backup'] is True
        assert 'backup_info' in result[0]
    
    def test_wps_metadata_extraction(self):
        """测试WPS元数据提取"""
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test.et')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        metadata = parser._extract_wps_metadata()
        
        # 验证元数据结构
        assert metadata['format_type'] == 'wps'
        assert metadata['creator'] == 'WPS Office'
        assert 'version' in metadata
        assert 'created_time' in metadata
        assert 'modified_time' in metadata
    
    def test_template_variables_extraction(self):
        """测试模板变量提取"""
        # 创建测试数据
        test_sheet = {
            'data': [
                ['姓名', '{{姓名}}', '部门'],
                ['{{员工姓名}}', '{{员工部门}}', '{{员工职位}}'],
                ['张三', '技术部', '工程师']
            ]
        }
        
        # 创建临时文件用于初始化解析器
        temp_file = os.path.join(self.temp_dir, 'dummy.ett')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        template_vars = parser._extract_template_variables(test_sheet)
        
        # 验证模板变量
        assert '姓名' in template_vars
        assert '员工姓名' in template_vars
        assert '员工部门' in template_vars
        assert '员工职位' in template_vars
        
        # 验证变量位置信息
        assert template_vars['姓名']['row'] == 0
        assert template_vars['姓名']['col'] == 1
    
    def test_backup_info_extraction(self):
        """测试备份信息提取"""
        # 创建临时文件
        temp_file = os.path.join(self.temp_dir, 'test_backup_20231201.ets')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        backup_info = parser._extract_backup_info()
        
        # 验证备份信息
        assert 'backup_time' in backup_info
        assert 'backup_reason' in backup_info
        assert backup_info['backup_reason'] == 'manual'  # 文件名包含backup
    
    def test_wps_format_detection(self):
        """测试WPS格式检测"""
        # 测试支持的WPS格式
        supported_formats = ['.et', '.ett', '.ets']
        for fmt in supported_formats:
            assert self.config.file_format.is_supported_format(fmt)
            assert self.config.file_format.get_format_type(fmt) == 'wps'
        
        # 测试不支持的格式
        unsupported_formats = ['.xyz', '.abc', '.123']
        for fmt in unsupported_formats:
            assert not self.config.file_format.is_supported_format(fmt)
            assert self.config.file_format.get_format_type(fmt) == 'unknown'

    @patch('mcp_sheet_parser.parser.openpyxl.load_workbook')
    def test_wps_workbook_annotations_links_merges_styles(self, mock_load_workbook):
        """测试WPS工作簿批注、超链接、合并单元格、样式提取"""
        mock_wb = MagicMock()
        mock_ws = MagicMock()
        mock_cell = MagicMock()
        # 模拟批注、超链接、样式
        mock_ws.title = 'Sheet1'
        mock_ws.max_row = 2
        mock_ws.max_column = 2
        # 单元格(1,1)有批注和超链接
        def cell_side_effect(row, column):
            cell = MagicMock()
            cell.value = f"R{row}C{column}"
            cell.data_type = 's'
            if row == 1 and column == 1:
                cell.comment = MagicMock()
                cell.comment.text = "批注内容"
                cell.hyperlink = MagicMock()
                cell.hyperlink.target = "https://example.com"
                cell.font = MagicMock()
                cell.font.bold = True
                cell.font.italic = False
                cell.font.sz = 14
                cell.font.name = "Arial"
                cell.font.underline = None
                cell.font.strike = False
                cell.font.color = MagicMock()
                cell.font.color.type = 'rgb'
                cell.font.color.rgb = 'FF0000'
                cell.fill = MagicMock()
                cell.fill.fgColor = MagicMock()
                cell.fill.fgColor.type = 'rgb'
                cell.fill.fgColor.rgb = 'FFFF00'
                cell.alignment = MagicMock()
                cell.alignment.horizontal = "center"
                cell.alignment.vertical = "middle"
                cell.alignment.wrap_text = True
                cell.border = MagicMock()
                for side in ['top', 'bottom', 'left', 'right']:
                    border_side = MagicMock()
                    border_side.style = "thin"
                    border_side.color = MagicMock()
                    border_side.color.rgb = '000000'
                    setattr(cell.border, side, border_side)
            else:
                cell.comment = None
                cell.hyperlink = None
                cell.font = MagicMock()
                cell.font.bold = False
                cell.font.italic = False
                cell.font.sz = 11
                cell.font.name = "Calibri"
                cell.font.underline = None
                cell.font.strike = False
                cell.font.color = None
                cell.fill = MagicMock()
                cell.fill.fgColor = MagicMock()
                cell.fill.fgColor.type = 'rgb'
                cell.fill.fgColor.rgb = 'FFFFFF'
                cell.alignment = MagicMock()
                cell.alignment.horizontal = "left"
                cell.alignment.vertical = "top"
                cell.alignment.wrap_text = False
                cell.border = MagicMock()
                for side in ['top', 'bottom', 'left', 'right']:
                    border_side = MagicMock()
                    border_side.style = None
                    border_side.color = None
                    setattr(cell.border, side, border_side)
            return cell
        mock_ws.cell.side_effect = cell_side_effect
        # 合并单元格
        mock_ws.merged_cells.ranges = [MagicMock(min_row=1, min_col=1, max_row=2, max_col=2)]
        mock_wb.worksheets = [mock_ws]
        mock_wb.sheetnames = ['Sheet1']
        mock_load_workbook.return_value = mock_wb
        temp_file = os.path.join(self.temp_dir, 'test.et')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        parser = SheetParser(temp_file, self.config)
        result = parser._parse_wps_workbook()
        assert len(result) == 1
        sheet = result[0]
        # 检查批注
        assert 'wps_comments' in sheet
        assert '0_0' in sheet['wps_comments']
        assert sheet['wps_comments']['0_0'] == "批注内容"
        # 检查超链接
        assert 'wps_hyperlinks' in sheet
        assert '0_0' in sheet['wps_hyperlinks']
        assert sheet['wps_hyperlinks']['0_0'] == "https://example.com"
        # 检查合并单元格
        assert 'wps_merged_cells' in sheet
        assert (0, 0, 1, 1) in sheet['wps_merged_cells']
        # 检查样式
        assert 'wps_styles' in sheet
        style = sheet['wps_styles'][0][0]
        assert style['bold'] is True
        assert style['font_size'] == 14
        assert style['font_name'] == "Arial"
        assert style['font_color'] == "#FF0000"
        assert style['bg_color'] == "#FFFF00"
        assert style['align'] == "center"
        assert style['valign'] == "middle"
        assert style['wrap_text'] is True
        assert 'border' in style
        for side in ['top', 'bottom', 'left', 'right']:
            assert style['border'][side]['style'] == "thin"
            assert style['border'][side]['color'] == "#000000"


if __name__ == '__main__':
    pytest.main([__file__]) 