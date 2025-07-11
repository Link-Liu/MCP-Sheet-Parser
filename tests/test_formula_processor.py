# test_formula_processor.py
# 公式处理模块测试

import unittest
from unittest.mock import Mock, patch

from mcp_sheet_parser.formula_processor import (
    FormulaProcessor, FormulaCalculator, CellReference,
    FormulaInfo, FormulaType, FormulaError
)
from mcp_sheet_parser.config import Config


class TestCellReference(unittest.TestCase):
    """单元格引用解析器测试"""
    
    def setUp(self):
        self.cell_ref = CellReference()
    
    def test_parse_simple_cell_reference(self):
        """测试简单单元格引用解析"""
        # 测试正常引用
        sheet, row, col = self.cell_ref.parse_cell_reference("A1")
        self.assertIsNone(sheet)
        self.assertEqual(row, 0)
        self.assertEqual(col, 0)
        
        # 测试其他单元格
        sheet, row, col = self.cell_ref.parse_cell_reference("B2")
        self.assertEqual(row, 1)
        self.assertEqual(col, 1)
        
        # 测试列字母转换
        sheet, row, col = self.cell_ref.parse_cell_reference("Z1")
        self.assertEqual(col, 25)
        
        # 测试双字母列
        sheet, row, col = self.cell_ref.parse_cell_reference("AA1")
        self.assertEqual(col, 26)
    
    def test_parse_absolute_reference(self):
        """测试绝对引用解析"""
        sheet, row, col = self.cell_ref.parse_cell_reference("$A$1")
        self.assertEqual(row, 0)
        self.assertEqual(col, 0)
        
        sheet, row, col = self.cell_ref.parse_cell_reference("$B2")
        self.assertEqual(row, 1)
        self.assertEqual(col, 1)
    
    def test_parse_sheet_reference(self):
        """测试工作表引用解析"""
        sheet, row, col = self.cell_ref.parse_cell_reference("Sheet1!A1")
        self.assertEqual(sheet, "Sheet1")
        self.assertEqual(row, 0)
        self.assertEqual(col, 0)
    
    def test_parse_range_reference(self):
        """测试范围引用解析"""
        sheet, start_row, start_col, end_row, end_col = self.cell_ref.parse_range_reference("A1:B2")
        self.assertIsNone(sheet)
        self.assertEqual(start_row, 0)
        self.assertEqual(start_col, 0)
        self.assertEqual(end_row, 1)
        self.assertEqual(end_col, 1)
        
        # 测试工作表范围引用
        sheet, start_row, start_col, end_row, end_col = self.cell_ref.parse_range_reference("Sheet1!A1:C3")
        self.assertEqual(sheet, "Sheet1")
        self.assertEqual(end_row, 2)
        self.assertEqual(end_col, 2)
    
    def test_invalid_references(self):
        """测试无效引用"""
        sheet, row, col = self.cell_ref.parse_cell_reference("invalid")
        self.assertEqual(row, -1)
        self.assertEqual(col, -1)
        
        sheet, start_row, start_col, end_row, end_col = self.cell_ref.parse_range_reference("invalid")
        self.assertEqual(start_row, -1)


class TestFormulaCalculator(unittest.TestCase):
    """公式计算引擎测试"""
    
    def setUp(self):
        # 创建测试数据
        self.sheet_data = {
            'data': [
                ['产品', '单价', '数量', '总价'],
                ['产品A', '100', '5', '=B2*C2'],
                ['产品B', '200', '3', '=B3*C3'],
                ['产品C', '150', '4', '=B4*C4'],
                ['合计', '', '', '=SUM(D2:D4)']
            ]
        }
        self.calculator = FormulaCalculator(self.sheet_data)
    
    def test_simple_math_calculation(self):
        """测试简单数学运算"""
        formula_info = self.calculator.calculate_formula("2+3")
        self.assertEqual(formula_info.calculated_value, 5)
        self.assertIsNone(formula_info.error)
        self.assertEqual(formula_info.formula_type, FormulaType.SIMPLE_MATH)
        
        # 测试乘法
        formula_info = self.calculator.calculate_formula("10*5")
        self.assertEqual(formula_info.calculated_value, 50)
        
        # 测试除法
        formula_info = self.calculator.calculate_formula("20/4")
        self.assertEqual(formula_info.calculated_value, 5)
        
        # 测试除零错误
        formula_info = self.calculator.calculate_formula("10/0")
        self.assertEqual(formula_info.error, FormulaError.DIV_ZERO)
    
    def test_cell_reference_calculation(self):
        """测试单元格引用计算"""
        # 测试有效引用
        formula_info = self.calculator.calculate_formula("B2")
        self.assertEqual(formula_info.calculated_value, 100)
        self.assertEqual(formula_info.formula_type, FormulaType.REFERENCE)
        
        # 测试字符串单元格
        formula_info = self.calculator.calculate_formula("A2")
        self.assertEqual(formula_info.calculated_value, "产品A")
        
        # 测试超出范围的引用
        formula_info = self.calculator.calculate_formula("Z99")
        self.assertEqual(formula_info.calculated_value, 0)  # 空单元格返回0
    
    def test_sum_function(self):
        """测试SUM函数"""
        # 测试简单数值求和
        formula_info = self.calculator.calculate_formula("SUM(1,2,3)")
        self.assertEqual(formula_info.calculated_value, 6)
        self.assertEqual(formula_info.formula_type, FormulaType.FUNCTION)
        
        # 测试单元格范围求和（模拟）
        # 注意：这里需要mock数据来测试范围
        with patch.object(self.calculator, '_resolve_range_reference') as mock_range:
            mock_range.return_value = [100, 200, 150]
            formula_info = self.calculator.calculate_formula("SUM(B2:B4)")
            self.assertEqual(formula_info.calculated_value, 450)
    
    def test_average_function(self):
        """测试AVERAGE函数"""
        formula_info = self.calculator.calculate_formula("AVERAGE(10,20,30)")
        self.assertEqual(formula_info.calculated_value, 20)
        
        # 测试单个值
        formula_info = self.calculator.calculate_formula("AVERAGE(42)")
        self.assertEqual(formula_info.calculated_value, 42)
    
    def test_count_function(self):
        """测试COUNT函数"""
        formula_info = self.calculator.calculate_formula("COUNT(1,2,3)")
        self.assertEqual(formula_info.calculated_value, 3)
        
        # 测试混合数据（只计算数值）
        with patch.object(self.calculator, '_parse_function_args') as mock_args:
            mock_args.return_value = [1, "text", 3, 4]
            formula_info = self.calculator.calculate_formula("COUNT(1,text,3,4)")
            self.assertEqual(formula_info.calculated_value, 3)
    
    def test_max_min_functions(self):
        """测试MAX和MIN函数"""
        formula_info = self.calculator.calculate_formula("MAX(10,5,20,15)")
        self.assertEqual(formula_info.calculated_value, 20)
        
        formula_info = self.calculator.calculate_formula("MIN(10,5,20,15)")
        self.assertEqual(formula_info.calculated_value, 5)
    
    def test_math_functions(self):
        """测试数学函数"""
        # ABS函数
        formula_info = self.calculator.calculate_formula("ABS(-5)")
        self.assertEqual(formula_info.calculated_value, 5)
        
        # ROUND函数
        formula_info = self.calculator.calculate_formula("ROUND(3.14159,2)")
        self.assertEqual(formula_info.calculated_value, 3.14)
        
        # INT函数
        formula_info = self.calculator.calculate_formula("INT(3.7)")
        self.assertEqual(formula_info.calculated_value, 3)
        
        # SQRT函数
        formula_info = self.calculator.calculate_formula("SQRT(16)")
        self.assertEqual(formula_info.calculated_value, 4)
        
        # SQRT负数错误
        formula_info = self.calculator.calculate_formula("SQRT(-1)")
        self.assertEqual(formula_info.error, FormulaError.NUM_ERROR)
    
    def test_logical_functions(self):
        """测试逻辑函数"""
        # IF函数
        formula_info = self.calculator.calculate_formula('IF(5>3,"大","小")')
        self.assertEqual(formula_info.calculated_value, "大")
        
        formula_info = self.calculator.calculate_formula('IF(2>5,"大","小")')
        self.assertEqual(formula_info.calculated_value, "小")
        
        # AND函数
        formula_info = self.calculator.calculate_formula("AND(1,1,1)")
        self.assertTrue(formula_info.calculated_value)
        
        formula_info = self.calculator.calculate_formula("AND(1,0,1)")
        self.assertFalse(formula_info.calculated_value)
        
        # OR函数
        formula_info = self.calculator.calculate_formula("OR(0,0,1)")
        self.assertTrue(formula_info.calculated_value)
        
        # NOT函数
        formula_info = self.calculator.calculate_formula("NOT(0)")
        self.assertTrue(formula_info.calculated_value)
    
    def test_string_functions(self):
        """测试字符串函数"""
        # LEN函数
        formula_info = self.calculator.calculate_formula('LEN("Hello")')
        self.assertEqual(formula_info.calculated_value, 5)
        
        # LEFT函数
        formula_info = self.calculator.calculate_formula('LEFT("Hello",3)')
        self.assertEqual(formula_info.calculated_value, "Hel")
        
        # RIGHT函数
        formula_info = self.calculator.calculate_formula('RIGHT("Hello",2)')
        self.assertEqual(formula_info.calculated_value, "lo")
        
        # MID函数
        formula_info = self.calculator.calculate_formula('MID("Hello",2,2)')
        self.assertEqual(formula_info.calculated_value, "el")
        
        # UPPER函数
        formula_info = self.calculator.calculate_formula('UPPER("hello")')
        self.assertEqual(formula_info.calculated_value, "HELLO")
        
        # LOWER函数
        formula_info = self.calculator.calculate_formula('LOWER("HELLO")')
        self.assertEqual(formula_info.calculated_value, "hello")
        
        # CONCATENATE函数
        formula_info = self.calculator.calculate_formula('CONCATENATE("A","B","C")')
        self.assertEqual(formula_info.calculated_value, "ABC")
    
    def test_error_handling(self):
        """测试错误处理"""
        # 不支持的函数
        formula_info = self.calculator.calculate_formula("UNKNOWN(1,2,3)")
        self.assertEqual(formula_info.error, FormulaError.NAME_ERROR)
        
        # 参数错误
        formula_info = self.calculator.calculate_formula("ABS()")
        self.assertEqual(formula_info.error, FormulaError.VALUE_ERROR)
        
        # 语法错误
        formula_info = self.calculator.calculate_formula("1++2")
        self.assertEqual(formula_info.error, FormulaError.VALUE_ERROR)
    
    def test_formula_type_detection(self):
        """测试公式类型检测"""
        # 简单数学运算
        formula_type = self.calculator._detect_formula_type("1+2*3")
        self.assertEqual(formula_type, FormulaType.SIMPLE_MATH)
        
        # 函数调用
        formula_type = self.calculator._detect_formula_type("SUM(A1:A10)")
        self.assertEqual(formula_type, FormulaType.FUNCTION)
        
        # 单元格引用
        formula_type = self.calculator._detect_formula_type("A1")
        self.assertEqual(formula_type, FormulaType.REFERENCE)
        
        # 条件公式
        formula_type = self.calculator._detect_formula_type("IF(A1>0,1,0)")
        self.assertEqual(formula_type, FormulaType.CONDITIONAL)
    
    def test_dependency_extraction(self):
        """测试依赖关系提取"""
        # 单个引用
        deps = self.calculator._extract_dependencies("A1+B2")
        self.assertIn("A1", deps)
        self.assertIn("B2", deps)
        
        # 范围引用
        deps = self.calculator._extract_dependencies("SUM(A1:A10)")
        self.assertIn("A1:A10", deps)
        
        # 工作表引用
        deps = self.calculator._extract_dependencies("Sheet1!A1+Sheet2!B2")
        self.assertIn("Sheet1!A1", deps)
        self.assertIn("Sheet2!B2", deps)


class TestFormulaProcessor(unittest.TestCase):
    """公式处理器测试"""
    
    def setUp(self):
        self.config = Config()
        self.processor = FormulaProcessor(self.config)
        
        # 创建包含公式的测试数据
        self.sheet_data = {
            'data': [
                ['A', 'B', 'C', 'D'],
                ['10', '20', '=A2+B2', '=C2*2'],
                ['15', '25', '=A3+B3', '=C3*2'],
                ['', '', '=SUM(C2:C3)', '=AVERAGE(D2:D3)']
            ],
            'styles': [
                [{}, {}, {}, {}],
                [{}, {}, {}, {}],
                [{}, {}, {}, {}],
                [{}, {}, {}, {}]
            ]
        }
    
    def test_process_sheet_formulas(self):
        """测试工作表公式处理"""
        enhanced_data = self.processor.process_sheet_formulas(self.sheet_data)
        
        # 检查是否添加了formulas字段
        self.assertIn('formulas', enhanced_data)
        
        formulas = enhanced_data['formulas']
        
        # 检查公式是否被正确识别
        self.assertIn('1_2', formulas)  # C2的公式
        self.assertIn('1_3', formulas)  # D2的公式
        self.assertIn('2_2', formulas)  # C3的公式
        self.assertIn('2_3', formulas)  # D3的公式
        self.assertIn('3_2', formulas)  # C4的公式
        self.assertIn('3_3', formulas)  # D4的公式
        
        # 检查公式信息
        formula_c2 = formulas['1_2']
        self.assertEqual(formula_c2.original_formula, '=A2+B2')
        self.assertEqual(formula_c2.calculated_value, 30)
        self.assertIsNone(formula_c2.error)
    
    def test_formula_detection(self):
        """测试公式检测"""
        self.assertTrue(self.processor.is_formula_cell('=SUM(A1:A10)'))
        self.assertTrue(self.processor.is_formula_cell('=A1+B1'))
        self.assertFalse(self.processor.is_formula_cell('123'))
        self.assertFalse(self.processor.is_formula_cell('text'))
        self.assertFalse(self.processor.is_formula_cell(''))
    
    def test_formula_caching(self):
        """测试公式缓存"""
        # 处理公式
        self.processor.process_sheet_formulas(self.sheet_data)
        
        # 检查缓存
        formula_info = self.processor.get_formula_info(1, 2)
        self.assertIsNotNone(formula_info)
        self.assertEqual(formula_info.original_formula, '=A2+B2')
        
        # 检查不存在的公式
        formula_info = self.processor.get_formula_info(0, 0)
        self.assertIsNone(formula_info)
    
    def test_statistics(self):
        """测试统计信息"""
        # 处理公式
        self.processor.process_sheet_formulas(self.sheet_data)
        
        # 检查统计信息
        stats = self.processor.get_formula_statistics()
        
        self.assertGreater(stats['total_formulas'], 0)
        self.assertGreaterEqual(stats['calculated_formulas'], 0)
        self.assertIn('function_usage', stats)
        self.assertIn('error_types', stats)
        
        # 检查函数使用统计
        if 'SUM' in stats['function_usage']:
            self.assertGreater(stats['function_usage']['SUM'], 0)


if __name__ == '__main__':
    unittest.main() 