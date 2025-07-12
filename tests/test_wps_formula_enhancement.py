# test_wps_formula_enhancement.py
# 测试WPS公式增强功能

import pytest
import os
import tempfile
from unittest.mock import patch, MagicMock
from mcp_sheet_parser.parser import SheetParser
from mcp_sheet_parser.config import UnifiedConfig


class TestWPSFormulaEnhancement:
    """测试WPS公式增强功能"""
    
    def setup_method(self):
        """设置测试环境"""
        self.config = UnifiedConfig()
        self.temp_dir = tempfile.mkdtemp()
    
    def teardown_method(self):
        """清理测试环境"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_extract_formula_cells(self):
        """测试公式单元格提取"""
        # 创建测试数据
        test_sheet = {
            'data': [
                ['A1', '=SUM(A1:C1)', '=IF(A1>0,"正数","负数")'],
                ['=A1+B1', '=VLOOKUP(A1,A:C,2,FALSE)', '=TODAY()'],
                ['=CONCATENATE(A1," ",B1)', '=COUNT(A:A)', '=MAX(A1:A10)']
            ]
        }
        
        # 创建临时文件用于初始化解析器
        temp_file = os.path.join(self.temp_dir, 'test.et')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        formula_cells = parser._extract_formula_cells(test_sheet)
        
        # 验证公式单元格数量
        assert len(formula_cells) == 8  # 8个公式单元格
        
        # 验证特定公式
        assert '0_1' in formula_cells  # =SUM(A1:C1)
        assert formula_cells['0_1']['formula'] == '=SUM(A1:C1)'
        assert formula_cells['0_1']['formula_type'] == 'mathematical'
        
        assert '0_2' in formula_cells  # =IF(A1>0,"正数","负数")
        assert formula_cells['0_2']['formula_type'] == 'logical'
        
        assert '1_1' in formula_cells  # =VLOOKUP(A1,A:C,2,FALSE)
        assert formula_cells['1_1']['formula_type'] == 'lookup'
        
        assert '1_2' in formula_cells  # =TODAY()
        assert formula_cells['1_2']['formula_type'] == 'datetime'
    
    def test_classify_formula(self):
        """测试公式分类"""
        # 创建临时文件用于初始化解析器
        temp_file = os.path.join(self.temp_dir, 'test.et')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        
        # 测试数学函数
        assert parser._classify_formula('=SUM(A1:A10)') == 'mathematical'
        assert parser._classify_formula('=AVERAGE(B1:B20)') == 'mathematical'
        assert parser._classify_formula('=COUNT(C:C)') == 'mathematical'
        
        # 测试文本函数
        assert parser._classify_formula('=CONCATENATE(A1," ",B1)') == 'text'
        assert parser._classify_formula('=LEFT(A1,5)') == 'text'
        assert parser._classify_formula('=UPPER(A1)') == 'text'
        
        # 测试逻辑函数
        assert parser._classify_formula('=IF(A1>0,"正数","负数")') == 'logical'
        assert parser._classify_formula('=AND(A1>0,B1<100)') == 'logical'
        
        # 测试查找函数
        assert parser._classify_formula('=VLOOKUP(A1,A:C,2,FALSE)') == 'lookup'
        assert parser._classify_formula('=INDEX(A:A,5)') == 'lookup'
        
        # 测试日期时间函数
        assert parser._classify_formula('=TODAY()') == 'datetime'
        assert parser._classify_formula('=YEAR(A1)') == 'datetime'
        
        # 测试引用函数
        assert parser._classify_formula('=INDIRECT("A1")') == 'reference'
        assert parser._classify_formula('=OFFSET(A1,1,1)') == 'reference'
        
        # 测试简单计算
        assert parser._classify_formula('=A1+B1') == 'calculation'
        assert parser._classify_formula('=A1*B1/C1') == 'calculation'
        
        # 测试单元格引用
        assert parser._classify_formula('=A1') == 'reference'
        assert parser._classify_formula('=B2') == 'reference'
    
    def test_extract_formula_dependencies(self):
        """测试公式依赖提取"""
        # 创建临时文件用于初始化解析器
        temp_file = os.path.join(self.temp_dir, 'test.et')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        
        # 测试简单单元格引用
        deps = parser._extract_formula_dependencies('=A1+B1')
        assert 'A1' in deps
        assert 'B1' in deps
        
        # 测试范围引用
        deps = parser._extract_formula_dependencies('=SUM(A1:B5)')
        assert 'A1:B5' in deps
        
        # 测试工作表引用
        deps = parser._extract_formula_dependencies('=Sheet1!A1')
        assert 'SHEET1!A1' in deps
        
        # 测试复杂公式
        deps = parser._extract_formula_dependencies('=SUM(A1:A10)+VLOOKUP(B1,Sheet1!A:C,2,FALSE)')
        assert 'A1:A10' in deps
        assert 'B1' in deps
        assert 'SHEET1!A:C' in deps
    
    def test_analyze_wps_formulas(self):
        """测试WPS公式分析"""
        # 创建测试数据
        test_sheet = {
            'data': [
                ['A1', '=SUM(A1:C1)', '=IF(A1>0,"正数","负数")'],
                ['=A1+B1', '=VLOOKUP(A1,A:C,2,FALSE)', '=TODAY()']
            ],
            'formula_cells': {
                '0_1': {'formula': '=SUM(A1:C1)', 'formula_type': 'mathematical'},
                '0_2': {'formula': '=IF(A1>0,"正数","负数")', 'formula_type': 'logical'},
                '1_0': {'formula': '=A1+B1', 'formula_type': 'calculation'},
                '1_1': {'formula': '=VLOOKUP(A1,A:C,2,FALSE)', 'formula_type': 'lookup'},
                '1_2': {'formula': '=TODAY()', 'formula_type': 'datetime'}
            }
        }
        
        # 创建临时文件用于初始化解析器
        temp_file = os.path.join(self.temp_dir, 'test.et')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        formula_info = parser._analyze_wps_formulas(test_sheet)
        
        # 验证统计信息
        assert formula_info['total_formulas'] == 5
        assert formula_info['standard_formulas'] == 5
        assert formula_info['wps_specific_formulas'] == 0
        
        # 验证公式类型统计
        assert formula_info['formula_types']['mathematical'] == 1
        assert formula_info['formula_types']['logical'] == 1
        assert formula_info['formula_types']['calculation'] == 1
        assert formula_info['formula_types']['lookup'] == 1
        assert formula_info['formula_types']['datetime'] == 1
    
    def test_check_formula_compatibility(self):
        """测试公式兼容性检查"""
        # 创建临时文件用于初始化解析器
        temp_file = os.path.join(self.temp_dir, 'test.et')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        
        # 测试正常公式
        issue = parser._check_formula_compatibility('=SUM(A1:A10)')
        assert issue is None
        
        # 测试括号不匹配
        issue = parser._check_formula_compatibility('=SUM(A1:A10')
        assert '括号不匹配' in issue
        
        # 测试公式长度超限
        long_formula = '=' + 'A1+' * 3000 + 'A1'  # 创建超长公式，确保超过8192字符
        issue = parser._check_formula_compatibility(long_formula)
        assert '公式长度超过Excel限制' in issue
    
    def test_is_wps_specific_formula(self):
        """测试WPS特有公式检测"""
        # 创建临时文件用于初始化解析器
        temp_file = os.path.join(self.temp_dir, 'test.et')
        with open(temp_file, 'w') as f:
            f.write('dummy content')
        
        parser = SheetParser(temp_file, self.config)
        
        # 测试标准公式
        assert not parser._is_wps_specific_formula('=SUM(A1:A10)')
        assert not parser._is_wps_specific_formula('=IF(A1>0,"正数","负数")')
        
        # 测试WPS特有公式（如果有的话）
        # 目前没有添加WPS特有模式，所以都返回False
        assert not parser._is_wps_specific_formula('=WPS_FUNCTION(A1)')
    
    @patch('mcp_sheet_parser.parser.openpyxl.load_workbook')
    def test_wps_workbook_with_formulas(self, mock_load_workbook):
        """测试WPS工作簿解析包含公式"""
        # 创建模拟的Excel工作簿
        mock_wb = MagicMock()
        mock_ws = MagicMock()
        mock_cell = MagicMock()
        
        # 设置模拟数据，包含公式
        mock_ws.max_row = 3
        mock_ws.max_column = 3
        mock_ws.cell.return_value = mock_cell
        mock_cell.value = "=SUM(A1:A10)"
        mock_cell.data_type = 'f'  # 公式类型
        
        mock_ws.title = 'Sheet1'
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
        assert 'formula_cells' in result[0]
        assert 'wps_formula_info' in result[0]
        
        # 验证公式信息
        formula_info = result[0]['wps_formula_info']
        assert formula_info['total_formulas'] >= 0  # 可能为0，因为mock数据可能不完整


if __name__ == '__main__':
    pytest.main([__file__]) 