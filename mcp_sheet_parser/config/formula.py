# formula.py
# 公式处理相关配置

from dataclasses import dataclass
from typing import Set, Dict, List


@dataclass
class FormulaConfig:
    """公式处理配置"""
    
    # 基本配置
    ENABLE_FORMULA_PROCESSING: bool = True
    SHOW_FORMULA_TEXT: bool = True
    CALCULATE_FORMULAS: bool = True
    SHOW_FORMULA_ERRORS: bool = True
    SUPPORTED_FUNCTIONS_ONLY: bool = False
    
    # 缓存配置
    FORMULA_CACHE_SIZE: int = 1000
    ENABLE_FORMULA_CACHING: bool = True
    CACHE_TIMEOUT_SECONDS: int = 3600
    
    # 计算配置
    MAX_CALCULATION_DEPTH: int = 100
    MAX_CIRCULAR_REFERENCES: int = 5
    CALCULATION_TIMEOUT_SECONDS: float = 5.0
    
    # 支持的函数列表
    SUPPORTED_FUNCTIONS: Set[str] = None
    
    # 函数分类
    FUNCTION_CATEGORIES: Dict[str, List[str]] = None
    
    def __post_init__(self):
        """初始化默认值"""
        if self.SUPPORTED_FUNCTIONS is None:
            self.SUPPORTED_FUNCTIONS = {
                # 数学函数
                'SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'ABS', 'ROUND',
                'SQRT', 'POWER', 'MOD', 'FLOOR', 'CEILING', 'INT',
                
                # 逻辑函数
                'IF', 'AND', 'OR', 'NOT', 'TRUE', 'FALSE',
                
                # 文本函数
                'CONCATENATE', 'LEFT', 'RIGHT', 'MID', 'LEN', 'TRIM',
                'UPPER', 'LOWER', 'SUBSTITUTE', 'FIND', 'SEARCH',
                
                # 日期函数
                'NOW', 'TODAY', 'DATE', 'TIME', 'YEAR', 'MONTH', 'DAY',
                'HOUR', 'MINUTE', 'SECOND', 'WEEKDAY',
                
                # 查找函数
                'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH',
                
                # 统计函数
                'COUNTIF', 'SUMIF', 'AVERAGEIF', 'MEDIAN', 'MODE',
                'STDEV', 'VAR'
            }
        
        if self.FUNCTION_CATEGORIES is None:
            self.FUNCTION_CATEGORIES = {
                'math': ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'ABS', 'ROUND', 'SQRT', 'POWER', 'MOD'],
                'logic': ['IF', 'AND', 'OR', 'NOT', 'TRUE', 'FALSE'],
                'text': ['CONCATENATE', 'LEFT', 'RIGHT', 'MID', 'LEN', 'TRIM', 'UPPER', 'LOWER'],
                'date': ['NOW', 'TODAY', 'DATE', 'TIME', 'YEAR', 'MONTH', 'DAY'],
                'lookup': ['VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH'],
                'statistical': ['COUNTIF', 'SUMIF', 'AVERAGEIF', 'MEDIAN', 'MODE', 'STDEV', 'VAR']
            }
    
    def is_function_supported(self, function_name: str) -> bool:
        """检查函数是否支持"""
        return function_name.upper() in self.SUPPORTED_FUNCTIONS
    
    def get_function_category(self, function_name: str) -> str:
        """获取函数分类"""
        func_upper = function_name.upper()
        for category, functions in self.FUNCTION_CATEGORIES.items():
            if func_upper in functions:
                return category
        return 'unknown'
    
    def get_supported_functions_by_category(self, category: str) -> List[str]:
        """根据分类获取支持的函数"""
        return self.FUNCTION_CATEGORIES.get(category, [])
    
    def should_process_formula(self, formula: str) -> bool:
        """判断是否应该处理公式"""
        if not self.ENABLE_FORMULA_PROCESSING:
            return False
        
        if not formula or not formula.startswith('='):
            return False
        
        if self.SUPPORTED_FUNCTIONS_ONLY:
            # 简单检查：提取函数名并验证
            import re
            functions = re.findall(r'([A-Z]+)\s*\(', formula.upper())
            return all(self.is_function_supported(func) for func in functions)
        
        return True
    
    def get_calculation_limits(self) -> Dict[str, any]:
        """获取计算限制配置"""
        return {
            'max_depth': self.MAX_CALCULATION_DEPTH,
            'max_circular_refs': self.MAX_CIRCULAR_REFERENCES,
            'timeout_seconds': self.CALCULATION_TIMEOUT_SECONDS
        }
    
    def get_cache_config(self) -> Dict[str, any]:
        """获取缓存配置"""
        return {
            'enabled': self.ENABLE_FORMULA_CACHING,
            'cache_size': self.FORMULA_CACHE_SIZE,
            'timeout_seconds': self.CACHE_TIMEOUT_SECONDS
        }
    
    def optimize_for_performance(self) -> 'FormulaConfig':
        """优化配置以提升性能"""
        optimized = FormulaConfig(
            ENABLE_FORMULA_PROCESSING=self.ENABLE_FORMULA_PROCESSING,
            SHOW_FORMULA_TEXT=self.SHOW_FORMULA_TEXT,
            CALCULATE_FORMULAS=False,  # 禁用计算以提升性能
            SHOW_FORMULA_ERRORS=False,
            SUPPORTED_FUNCTIONS_ONLY=True,  # 只处理支持的函数
            FORMULA_CACHE_SIZE=self.FORMULA_CACHE_SIZE * 2,  # 增大缓存
            ENABLE_FORMULA_CACHING=True,
            CACHE_TIMEOUT_SECONDS=self.CACHE_TIMEOUT_SECONDS,
            MAX_CALCULATION_DEPTH=50,  # 减少计算深度
            MAX_CIRCULAR_REFERENCES=2,
            CALCULATION_TIMEOUT_SECONDS=2.0,  # 减少超时时间
            SUPPORTED_FUNCTIONS=self.SUPPORTED_FUNCTIONS,
            FUNCTION_CATEGORIES=self.FUNCTION_CATEGORIES
        )
        
        return optimized
    
    def optimize_for_accuracy(self) -> 'FormulaConfig':
        """优化配置以提升准确性"""
        optimized = FormulaConfig(
            ENABLE_FORMULA_PROCESSING=True,
            SHOW_FORMULA_TEXT=True,
            CALCULATE_FORMULAS=True,
            SHOW_FORMULA_ERRORS=True,
            SUPPORTED_FUNCTIONS_ONLY=False,  # 尝试处理所有函数
            FORMULA_CACHE_SIZE=self.FORMULA_CACHE_SIZE,
            ENABLE_FORMULA_CACHING=True,
            CACHE_TIMEOUT_SECONDS=self.CACHE_TIMEOUT_SECONDS,
            MAX_CALCULATION_DEPTH=200,  # 增加计算深度
            MAX_CIRCULAR_REFERENCES=10,
            CALCULATION_TIMEOUT_SECONDS=10.0,  # 增加超时时间
            SUPPORTED_FUNCTIONS=self.SUPPORTED_FUNCTIONS,
            FUNCTION_CATEGORIES=self.FUNCTION_CATEGORIES
        )
        
        return optimized
    
    def get_display_options(self) -> Dict[str, bool]:
        """获取显示选项"""
        return {
            'show_formula_text': self.SHOW_FORMULA_TEXT,
            'show_formula_errors': self.SHOW_FORMULA_ERRORS,
            'calculate_formulas': self.CALCULATE_FORMULAS
        } 