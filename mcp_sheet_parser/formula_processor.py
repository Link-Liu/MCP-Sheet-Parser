# formula_processor.py
# 公式处理模块 - 公式识别、计算与错误处理

import re
import logging
import math
from typing import Dict, List, Any, Optional, Union, Tuple
from dataclasses import dataclass
from enum import Enum


class FormulaError(Enum):
    """公式错误类型"""
    REF_ERROR = "#REF!"      # 引用错误
    DIV_ZERO = "#DIV/0!"     # 除零错误
    VALUE_ERROR = "#VALUE!"   # 值错误
    NAME_ERROR = "#NAME?"     # 名称错误
    NUM_ERROR = "#NUM!"       # 数值错误
    NA_ERROR = "#N/A"         # 不可用错误
    NULL_ERROR = "#NULL!"     # 空值错误


class FormulaType(Enum):
    """公式类型"""
    SIMPLE_MATH = "simple_math"       # 简单数学运算
    FUNCTION = "function"             # 函数调用
    REFERENCE = "reference"           # 单元格引用
    ARRAY = "array"                   # 数组公式
    CONDITIONAL = "conditional"       # 条件公式
    COMPLEX = "complex"               # 复杂公式


@dataclass
class FormulaInfo:
    """公式信息"""
    original_formula: str           # 原始公式文本
    formula_type: FormulaType      # 公式类型
    calculated_value: Any          # 计算结果
    error: Optional[FormulaError]  # 错误信息
    dependencies: List[str]        # 依赖的单元格
    is_calculated: bool = False    # 是否已计算
    description: str = ""          # 公式描述


class CellReference:
    """单元格引用解析器"""
    
    def __init__(self):
        # 单元格引用正则模式 (如 A1, B2, $A$1, Sheet1!A1)
        self.cell_pattern = re.compile(
            r'(?:([^!]+)!)?\$?([A-Z]+)\$?(\d+)',
            re.IGNORECASE
        )
        
        # 范围引用正则模式 (如 A1:B10, $A$1:$B$10)
        self.range_pattern = re.compile(
            r'(?:([^!]+)!)?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)',
            re.IGNORECASE
        )
    
    def parse_cell_reference(self, ref: str) -> Tuple[Optional[str], int, int]:
        """
        解析单元格引用
        
        Args:
            ref: 单元格引用字符串 (如 "A1", "Sheet1!B2")
            
        Returns:
            (工作表名, 行号, 列号) 元组，索引从0开始
        """
        match = self.cell_pattern.match(ref.strip())
        if not match:
            return None, -1, -1
        
        sheet_name, col_letters, row_str = match.groups()
        
        try:
            row = int(row_str) - 1  # 转换为0-based索引
            col = self._letters_to_col(col_letters)
            return sheet_name, row, col
        except (ValueError, TypeError):
            return None, -1, -1
    
    def parse_range_reference(self, ref: str) -> Tuple[Optional[str], int, int, int, int]:
        """
        解析范围引用
        
        Args:
            ref: 范围引用字符串 (如 "A1:B10", "Sheet1!A1:B10")
            
        Returns:
            (工作表名, 起始行, 起始列, 结束行, 结束列) 元组
        """
        match = self.range_pattern.match(ref.strip())
        if not match:
            return None, -1, -1, -1, -1
        
        sheet_name, start_col_letters, start_row_str, end_col_letters, end_row_str = match.groups()
        
        try:
            start_row = int(start_row_str) - 1
            start_col = self._letters_to_col(start_col_letters)
            end_row = int(end_row_str) - 1
            end_col = self._letters_to_col(end_col_letters)
            
            return sheet_name, start_row, start_col, end_row, end_col
        except (ValueError, TypeError):
            return None, -1, -1, -1, -1
    
    def _letters_to_col(self, letters: str) -> int:
        """将列字母转换为列索引 (A=0, B=1, ..., Z=25, AA=26, ...)"""
        result = 0
        for char in letters.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1


class FormulaCalculator:
    """公式计算引擎"""
    
    def __init__(self, sheet_data: Dict[str, Any]):
        self.sheet_data = sheet_data
        self.cell_ref = CellReference()
        self.logger = logging.getLogger(__name__)
        
        # 支持的函数
        self.supported_functions = {
            'SUM': self._function_sum,
            'AVERAGE': self._function_average,
            'COUNT': self._function_count,
            'COUNTA': self._function_counta,
            'MAX': self._function_max,
            'MIN': self._function_min,
            'ABS': self._function_abs,
            'ROUND': self._function_round,
            'INT': self._function_int,
            'SQRT': self._function_sqrt,
            'POWER': self._function_power,
            'IF': self._function_if,
            'AND': self._function_and,
            'OR': self._function_or,
            'NOT': self._function_not,
            'LEN': self._function_len,
            'LEFT': self._function_left,
            'RIGHT': self._function_right,
            'MID': self._function_mid,
            'UPPER': self._function_upper,
            'LOWER': self._function_lower,
            'CONCATENATE': self._function_concatenate
        }
        
        # 数学运算符
        self.operators = {
            '+': lambda a, b: a + b,
            '-': lambda a, b: a - b,
            '*': lambda a, b: a * b,
            '/': lambda a, b: a / b if b != 0 else FormulaError.DIV_ZERO,
            '^': lambda a, b: a ** b,
            '%': lambda a, b: a % b
        }
    
    def calculate_formula(self, formula: str, current_row: int = 0, current_col: int = 0) -> FormulaInfo:
        """
        计算公式
        
        Args:
            formula: 公式字符串 (不包含等号)
            current_row: 当前单元格行号
            current_col: 当前单元格列号
            
        Returns:
            FormulaInfo对象
        """
        try:
            # 清理公式
            formula = formula.strip()
            if formula.startswith('='):
                formula = formula[1:]
            
            # 检测公式类型
            formula_type = self._detect_formula_type(formula)
            
            # 提取依赖关系
            dependencies = self._extract_dependencies(formula)
            
            # 计算结果
            result, error = self._evaluate_formula(formula, current_row, current_col)
            
            return FormulaInfo(
                original_formula=f"={formula}",
                formula_type=formula_type,
                calculated_value=result,
                error=error,
                dependencies=dependencies,
                is_calculated=True,
                description=self._generate_description(formula, formula_type)
            )
            
        except Exception as e:
            self.logger.error(f"公式计算错误: {formula}, 错误: {e}")
            return FormulaInfo(
                original_formula=f"={formula}",
                formula_type=FormulaType.COMPLEX,
                calculated_value=None,
                error=FormulaError.VALUE_ERROR,
                dependencies=[],
                is_calculated=False,
                description="计算错误"
            )
    
    def _detect_formula_type(self, formula: str) -> FormulaType:
        """检测公式类型"""
        # 检查是否为条件公式 (在函数检测之前，因为IF也是函数)
        if 'IF(' in formula.upper():
            return FormulaType.CONDITIONAL
        
        # 检查是否为数组公式
        if '{' in formula and '}' in formula:
            return FormulaType.ARRAY
        
        # 检查是否为函数调用
        if re.search(r'[A-Z]+\s*\(', formula, re.IGNORECASE):
            return FormulaType.FUNCTION
        
        # 检查是否为简单数学运算
        if re.match(r'^[\d\s+\-*/().\^%]+$', formula):
            return FormulaType.SIMPLE_MATH
        
        # 检查是否为单元格引用
        if re.match(r'^[A-Z]+\d+$', formula, re.IGNORECASE):
            return FormulaType.REFERENCE
        
        return FormulaType.COMPLEX
    
    def _extract_dependencies(self, formula: str) -> List[str]:
        """提取公式中的单元格依赖"""
        dependencies = []
        
        # 查找范围引用 (必须在单元格引用之前，避免被单独匹配)
        range_refs = re.findall(
            r'(?:[^!\s]+!)?[A-Z]+\d+:[A-Z]+\d+',
            formula,
            re.IGNORECASE
        )
        dependencies.extend(range_refs)
        
        # 查找单元格引用 (排除已经在范围中的)
        cell_refs = re.findall(
            r'(?:[^!\s]+!)?[A-Z]+\d+',
            formula,
            re.IGNORECASE
        )
        
        # 过滤掉已经在范围引用中的单元格引用
        for cell_ref in cell_refs:
            is_in_range = False
            for range_ref in range_refs:
                if cell_ref in range_ref:
                    is_in_range = True
                    break
            if not is_in_range:
                dependencies.append(cell_ref)
        
        return list(set(dependencies))  # 去重
    
    def _evaluate_formula(self, formula: str, current_row: int, current_col: int) -> Tuple[Any, Optional[FormulaError]]:
        """评估公式"""
        try:
            # 首先检查是否为错误值
            if formula.upper() in [e.value for e in FormulaError]:
                return formula, FormulaError(formula.upper())
            
            # 简单数学运算
            if re.match(r'^[\d\s+\-*/().\^%]+$', formula):
                # 检查语法错误（如连续的操作符）
                if re.search(r'[\+\-\*/\^%]{2,}', formula) or re.search(r'^[\+\*/\^%]|[\+\-\*/\^%]$', formula):
                    return FormulaError.VALUE_ERROR.value, FormulaError.VALUE_ERROR
                
                # 替换^为**
                formula = formula.replace('^', '**')
                try:
                    # 验证表达式语法
                    compile(formula, '<string>', 'eval')
                    result = eval(formula)
                    return result, None
                except ZeroDivisionError:
                    return FormulaError.DIV_ZERO.value, FormulaError.DIV_ZERO
                except (SyntaxError, ValueError):
                    return FormulaError.VALUE_ERROR.value, FormulaError.VALUE_ERROR
                except Exception:
                    return FormulaError.VALUE_ERROR.value, FormulaError.VALUE_ERROR
            
            # 单元格引用
            if re.match(r'^[A-Z]+\d+$', formula, re.IGNORECASE):
                return self._resolve_cell_reference(formula)
            
            # 函数调用
            if re.search(r'[A-Z]+\s*\(', formula, re.IGNORECASE):
                return self._evaluate_function(formula)
            
            # 复杂公式 - 逐步替换引用
            evaluated_formula = self._substitute_references(formula)
            
            # 尝试计算
            try:
                result = eval(evaluated_formula)
                return result, None
            except Exception:
                return FormulaError.VALUE_ERROR.value, FormulaError.VALUE_ERROR
                
        except Exception as e:
            self.logger.error(f"公式评估错误: {formula}, 错误: {e}")
            return FormulaError.VALUE_ERROR.value, FormulaError.VALUE_ERROR
    
    def _resolve_cell_reference(self, ref: str) -> Tuple[Any, Optional[FormulaError]]:
        """解析单元格引用"""
        result = self.cell_ref.parse_cell_reference(ref)
        if len(result) != 3:
            return FormulaError.REF_ERROR.value, FormulaError.REF_ERROR
            
        sheet_name, row, col = result
        
        if row == -1 or col == -1:
            return FormulaError.REF_ERROR.value, FormulaError.REF_ERROR
        
        # 获取单元格值
        try:
            data = self.sheet_data.get('data', [])
            if row >= len(data) or col >= len(data[row]):
                return 0, None  # 空单元格返回0
            
            value = data[row][col]
            
            # 尝试转换为数值
            if isinstance(value, (int, float)):
                return value, None
            elif isinstance(value, str):
                try:
                    if '.' in value:
                        return float(value), None
                    else:
                        return int(value), None
                except ValueError:
                    return value, None
            else:
                return value or 0, None
                
        except Exception:
            return FormulaError.REF_ERROR.value, FormulaError.REF_ERROR
    
    def _evaluate_function(self, formula: str) -> Tuple[Any, Optional[FormulaError]]:
        """评估函数"""
        # 提取函数名和参数
        match = re.match(r'([A-Z]+)\s*\((.*)\)', formula, re.IGNORECASE | re.DOTALL)
        if not match:
            return FormulaError.NAME_ERROR.value, FormulaError.NAME_ERROR
        
        func_name = match.group(1).upper()
        args_str = match.group(2)
        
        if func_name not in self.supported_functions:
            return FormulaError.NAME_ERROR.value, FormulaError.NAME_ERROR
        
        # 解析参数
        try:
            args = self._parse_function_args(args_str)
            func = self.supported_functions[func_name]
            result = func(args)
            
            if isinstance(result, FormulaError):
                return result.value, result
            else:
                return result, None
                
        except Exception as e:
            self.logger.error(f"函数计算错误: {func_name}, 错误: {e}")
            return FormulaError.VALUE_ERROR.value, FormulaError.VALUE_ERROR
    
    def _parse_function_args(self, args_str: str) -> List[Any]:
        """解析函数参数"""
        if not args_str.strip():
            return []
        
        args = []
        current_arg = ""
        paren_level = 0
        
        for char in args_str:
            if char == ',' and paren_level == 0:
                args.append(self._evaluate_arg(current_arg.strip()))
                current_arg = ""
            else:
                if char == '(':
                    paren_level += 1
                elif char == ')':
                    paren_level -= 1
                current_arg += char
        
        if current_arg.strip():
            args.append(self._evaluate_arg(current_arg.strip()))
        
        return args
    
    def _evaluate_arg(self, arg: str) -> Any:
        """评估单个参数"""
        if not arg:
            return 0
        
        # 字符串字面量
        if arg.startswith('"') and arg.endswith('"'):
            return arg[1:-1]
        
        # 数值
        try:
            if '.' in arg:
                return float(arg)
            else:
                return int(arg)
        except ValueError:
            pass
        
        # 单元格引用
        if re.match(r'^[A-Z]+\d+$', arg, re.IGNORECASE):
            value, error = self._resolve_cell_reference(arg)
            return value if error is None else 0
        
        # 范围引用
        if ':' in arg:
            return self._resolve_range_reference(arg)
        
        # 其他表达式
        try:
            return eval(arg)
        except:
            return arg
    
    def _resolve_range_reference(self, ref: str) -> List[Any]:
        """解析范围引用"""
        sheet_name, start_row, start_col, end_row, end_col = self.cell_ref.parse_range_reference(ref)
        
        if start_row == -1:
            return []
        
        values = []
        data = self.sheet_data.get('data', [])
        
        # 添加递归深度保护
        if not hasattr(self, '_recursion_depth'):
            self._recursion_depth = 0
        
        if self._recursion_depth > 10:  # 限制递归深度
            return []
        
        for row in range(start_row, min(end_row + 1, len(data))):
            for col in range(start_col, min(end_col + 1, len(data[row]) if row < len(data) else 0)):
                try:
                    value = data[row][col]
                    if isinstance(value, (int, float)):
                        values.append(value)
                    elif isinstance(value, str):
                        # 检查是否为公式
                        if value.startswith('='):
                            # 为了避免复杂的递归，先尝试简单的计算
                            # 对于简单的 =A1*B1 这样的公式，直接计算
                            try:
                                self._recursion_depth += 1
                                # 替换单元格引用为实际值
                                evaluated_formula = self._substitute_references(value[1:])
                                result = eval(evaluated_formula)
                                if isinstance(result, (int, float)):
                                    values.append(result)
                                self._recursion_depth -= 1
                            except:
                                self._recursion_depth -= 1
                                continue
                        else:
                            # 尝试转换为数值
                            try:
                                values.append(float(value) if '.' in value else int(value))
                            except ValueError:
                                continue
                except IndexError:
                    continue
        
        return values
    
    def _substitute_references(self, formula: str) -> str:
        """替换公式中的单元格引用"""
        # 替换单元格引用
        def replace_cell_ref(match):
            ref = match.group(0)
            value, error = self._resolve_cell_reference(ref)
            if error is None and isinstance(value, (int, float)):
                return str(value)
            return '0'
        
        formula = re.sub(
            r'[A-Z]+\d+',
            replace_cell_ref,
            formula,
            flags=re.IGNORECASE
        )
        
        return formula.replace('^', '**')
    
    def _generate_description(self, formula: str, formula_type: FormulaType) -> str:
        """生成公式描述"""
        descriptions = {
            FormulaType.SIMPLE_MATH: "数学运算",
            FormulaType.FUNCTION: "函数调用",
            FormulaType.REFERENCE: "单元格引用",
            FormulaType.ARRAY: "数组公式",
            FormulaType.CONDITIONAL: "条件公式",
            FormulaType.COMPLEX: "复杂公式"
        }
        
        base_desc = descriptions.get(formula_type, "未知公式")
        
        # 尝试识别主要函数
        if formula_type == FormulaType.FUNCTION:
            match = re.search(r'([A-Z]+)\s*\(', formula, re.IGNORECASE)
            if match:
                func_name = match.group(1).upper()
                return f"{func_name}函数"
        
        return base_desc
    
    # ========== 内置函数实现 ==========
    
    def _function_sum(self, args: List[Any]) -> Union[float, FormulaError]:
        """SUM函数"""
        try:
            total = 0
            for arg in args:
                if isinstance(arg, list):
                    total += sum(v for v in arg if isinstance(v, (int, float)))
                elif isinstance(arg, (int, float)):
                    total += arg
            return total
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_average(self, args: List[Any]) -> Union[float, FormulaError]:
        """AVERAGE函数"""
        try:
            values = []
            for arg in args:
                if isinstance(arg, list):
                    values.extend(v for v in arg if isinstance(v, (int, float)))
                elif isinstance(arg, (int, float)):
                    values.append(arg)
            
            return sum(values) / len(values) if values else 0
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_count(self, args: List[Any]) -> Union[int, FormulaError]:
        """COUNT函数 - 计算数值个数"""
        try:
            count = 0
            for arg in args:
                if isinstance(arg, list):
                    count += sum(1 for v in arg if isinstance(v, (int, float)))
                elif isinstance(arg, (int, float)):
                    count += 1
            return count
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_counta(self, args: List[Any]) -> Union[int, FormulaError]:
        """COUNTA函数 - 计算非空值个数"""
        try:
            count = 0
            for arg in args:
                if isinstance(arg, list):
                    count += sum(1 for v in arg if v is not None and v != "")
                elif arg is not None and arg != "":
                    count += 1
            return count
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_max(self, args: List[Any]) -> Union[float, FormulaError]:
        """MAX函数"""
        try:
            values = []
            for arg in args:
                if isinstance(arg, list):
                    values.extend(v for v in arg if isinstance(v, (int, float)))
                elif isinstance(arg, (int, float)):
                    values.append(arg)
            
            return max(values) if values else 0
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_min(self, args: List[Any]) -> Union[float, FormulaError]:
        """MIN函数"""
        try:
            values = []
            for arg in args:
                if isinstance(arg, list):
                    values.extend(v for v in arg if isinstance(v, (int, float)))
                elif isinstance(arg, (int, float)):
                    values.append(arg)
            
            return min(values) if values else 0
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_abs(self, args: List[Any]) -> Union[float, FormulaError]:
        """ABS函数"""
        if len(args) != 1:
            return FormulaError.VALUE_ERROR
        
        try:
            value = args[0]
            if isinstance(value, (int, float)):
                return abs(value)
            return FormulaError.VALUE_ERROR
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_round(self, args: List[Any]) -> Union[float, FormulaError]:
        """ROUND函数"""
        if len(args) < 1 or len(args) > 2:
            return FormulaError.VALUE_ERROR
        
        try:
            value = args[0]
            digits = args[1] if len(args) > 1 else 0
            
            if isinstance(value, (int, float)) and isinstance(digits, (int, float)):
                return round(value, int(digits))
            return FormulaError.VALUE_ERROR
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_int(self, args: List[Any]) -> Union[int, FormulaError]:
        """INT函数"""
        if len(args) != 1:
            return FormulaError.VALUE_ERROR
        
        try:
            value = args[0]
            if isinstance(value, (int, float)):
                return int(value)
            return FormulaError.VALUE_ERROR
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_sqrt(self, args: List[Any]) -> Union[float, FormulaError]:
        """SQRT函数"""
        if len(args) != 1:
            return FormulaError.VALUE_ERROR
        
        try:
            value = args[0]
            if isinstance(value, (int, float)):
                if value < 0:
                    return FormulaError.NUM_ERROR
                return math.sqrt(value)
            return FormulaError.VALUE_ERROR
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_power(self, args: List[Any]) -> Union[float, FormulaError]:
        """POWER函数"""
        if len(args) != 2:
            return FormulaError.VALUE_ERROR
        
        try:
            base, exponent = args
            if isinstance(base, (int, float)) and isinstance(exponent, (int, float)):
                return base ** exponent
            return FormulaError.VALUE_ERROR
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_if(self, args: List[Any]) -> Union[Any, FormulaError]:
        """IF函数"""
        if len(args) < 2 or len(args) > 3:
            return FormulaError.VALUE_ERROR
        
        try:
            condition = args[0]
            true_value = args[1]
            false_value = args[2] if len(args) > 2 else False
            
            # 简单的条件判断
            if condition:
                return true_value
            else:
                return false_value
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_and(self, args: List[Any]) -> Union[bool, FormulaError]:
        """AND函数"""
        try:
            return all(bool(arg) for arg in args)
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_or(self, args: List[Any]) -> Union[bool, FormulaError]:
        """OR函数"""
        try:
            return any(bool(arg) for arg in args)
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_not(self, args: List[Any]) -> Union[bool, FormulaError]:
        """NOT函数"""
        if len(args) != 1:
            return FormulaError.VALUE_ERROR
        
        try:
            return not bool(args[0])
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_len(self, args: List[Any]) -> Union[int, FormulaError]:
        """LEN函数"""
        if len(args) != 1:
            return FormulaError.VALUE_ERROR
        
        try:
            value = args[0]
            return len(str(value))
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_left(self, args: List[Any]) -> Union[str, FormulaError]:
        """LEFT函数"""
        if len(args) < 1 or len(args) > 2:
            return FormulaError.VALUE_ERROR
        
        try:
            text = str(args[0])
            num_chars = int(args[1]) if len(args) > 1 else 1
            return text[:num_chars]
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_right(self, args: List[Any]) -> Union[str, FormulaError]:
        """RIGHT函数"""
        if len(args) < 1 or len(args) > 2:
            return FormulaError.VALUE_ERROR
        
        try:
            text = str(args[0])
            num_chars = int(args[1]) if len(args) > 1 else 1
            return text[-num_chars:]
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_mid(self, args: List[Any]) -> Union[str, FormulaError]:
        """MID函数"""
        if len(args) != 3:
            return FormulaError.VALUE_ERROR
        
        try:
            text = str(args[0])
            start = int(args[1]) - 1  # Excel uses 1-based indexing
            length = int(args[2])
            return text[start:start + length]
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_upper(self, args: List[Any]) -> Union[str, FormulaError]:
        """UPPER函数"""
        if len(args) != 1:
            return FormulaError.VALUE_ERROR
        
        try:
            return str(args[0]).upper()
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_lower(self, args: List[Any]) -> Union[str, FormulaError]:
        """LOWER函数"""
        if len(args) != 1:
            return FormulaError.VALUE_ERROR
        
        try:
            return str(args[0]).lower()
        except Exception:
            return FormulaError.VALUE_ERROR
    
    def _function_concatenate(self, args: List[Any]) -> Union[str, FormulaError]:
        """CONCATENATE函数"""
        try:
            return ''.join(str(arg) for arg in args)
        except Exception:
            return FormulaError.VALUE_ERROR


class FormulaProcessor:
    """公式处理器主类"""
    
    def __init__(self, config=None):
        from .config import Config
        self.config = config or Config()
        self.logger = logging.getLogger(__name__)
        
        # 公式缓存
        self.formula_cache: Dict[str, FormulaInfo] = {}
        
        # 统计信息
        self.stats = {
            'total_formulas': 0,
            'calculated_formulas': 0,
            'error_formulas': 0,
            'function_usage': {},
            'error_types': {}
        }
    
    def process_sheet_formulas(self, sheet_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        处理工作表中的所有公式
        
        Args:
            sheet_data: 工作表数据
            
        Returns:
            包含公式信息的增强工作表数据
        """
        self.logger.info("开始处理工作表公式...")
        
        enhanced_data = sheet_data.copy()
        enhanced_data['formulas'] = {}
        
        calculator = FormulaCalculator(sheet_data)
        data = sheet_data.get('data', [])
        
        for row_idx, row in enumerate(data):
            for col_idx, cell_value in enumerate(row):
                if isinstance(cell_value, str) and cell_value.startswith('='):
                    # 发现公式
                    formula_key = f"{row_idx}_{col_idx}"
                    
                    # 计算公式
                    formula_info = calculator.calculate_formula(
                        cell_value, row_idx, col_idx
                    )
                    
                    # 存储公式信息
                    enhanced_data['formulas'][formula_key] = formula_info
                    
                    # 更新统计
                    self._update_statistics(formula_info)
                    
                    # 缓存公式
                    self.formula_cache[formula_key] = formula_info
                    
                    self.logger.debug(f"处理公式 [{row_idx},{col_idx}]: {cell_value}")
        
        self.logger.info(f"公式处理完成：{self.stats['total_formulas']}个公式，"
                        f"{self.stats['calculated_formulas']}个成功计算")
        
        return enhanced_data
    
    def get_formula_info(self, row: int, col: int) -> Optional[FormulaInfo]:
        """获取指定位置的公式信息"""
        formula_key = f"{row}_{col}"
        return self.formula_cache.get(formula_key)
    
    def is_formula_cell(self, value: Any) -> bool:
        """判断是否为公式单元格"""
        return isinstance(value, str) and value.startswith('=')
    
    def get_formula_statistics(self) -> Dict[str, Any]:
        """获取公式统计信息"""
        return self.stats.copy()
    
    def _update_statistics(self, formula_info: FormulaInfo):
        """更新统计信息"""
        self.stats['total_formulas'] += 1
        
        if formula_info.is_calculated and formula_info.error is None:
            self.stats['calculated_formulas'] += 1
        else:
            self.stats['error_formulas'] += 1
        
        # 统计函数使用
        if formula_info.formula_type == FormulaType.FUNCTION:
            # 提取函数名
            match = re.search(r'([A-Z]+)\s*\(', formula_info.original_formula, re.IGNORECASE)
            if match:
                func_name = match.group(1).upper()
                self.stats['function_usage'][func_name] = self.stats['function_usage'].get(func_name, 0) + 1
        
        # 统计错误类型
        if formula_info.error:
            error_name = formula_info.error.name
            self.stats['error_types'][error_name] = self.stats['error_types'].get(error_name, 0) + 1 