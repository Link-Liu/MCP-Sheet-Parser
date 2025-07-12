# test_config_refactor.py
# 测试配置管理重构

import unittest
import tempfile
import os

# 从新的config子模块导入
from mcp_sheet_parser.config import (
    Config, UnifiedConfig,
    FileFormatConfig, PerformanceConfig, HTMLConfig, StyleConfig,
    FormulaConfig, ChartConfig, LoggingConfig
)

# 从原config.py模块导入工厂函数和常量
from mcp_sheet_parser import config as legacy_config

import logging


class TestFileFormatConfig(unittest.TestCase):
    """测试文件格式配置"""
    
    def test_default_formats(self):
        """测试默认文件格式"""
        config = FileFormatConfig()
        
        # 检查默认格式
        self.assertIn('.xlsx', config.SUPPORTED_EXCEL_FORMATS)
        self.assertIn('.csv', config.SUPPORTED_CSV_FORMATS)
        self.assertIn('.et', config.SUPPORTED_WPS_FORMATS)
    
    def test_format_checking(self):
        """测试格式检查方法"""
        config = FileFormatConfig()
        
        self.assertTrue(config.is_excel_format('.xlsx'))
        self.assertTrue(config.is_csv_format('.csv'))
        self.assertTrue(config.is_wps_format('.et'))
        self.assertFalse(config.is_excel_format('.txt'))
    
    def test_format_type_detection(self):
        """测试格式类型检测"""
        config = FileFormatConfig()
        
        self.assertEqual(config.get_format_type('.xlsx'), 'excel')
        self.assertEqual(config.get_format_type('.csv'), 'csv')
        self.assertEqual(config.get_format_type('.et'), 'wps')
        self.assertEqual(config.get_format_type('.txt'), 'unknown')


class TestPerformanceConfig(unittest.TestCase):
    """测试性能配置"""
    
    def test_default_values(self):
        """测试默认值"""
        config = PerformanceConfig()
        
        self.assertEqual(config.MAX_FILE_SIZE_MB, 500)
        self.assertEqual(config.MAX_ROWS, 1000000)
        self.assertEqual(config.CHUNK_SIZE, 1000)
        self.assertTrue(config.ENABLE_PERFORMANCE_MODE)
    
    def test_large_file_detection(self):
        """测试大文件检测"""
        config = PerformanceConfig()
        
        self.assertTrue(config.is_large_file(200))
        self.assertFalse(config.is_large_file(50))
    
    def test_streaming_decision(self):
        """测试流式处理决策"""
        config = PerformanceConfig()
        
        self.assertTrue(config.should_use_streaming(100))
        self.assertFalse(config.should_use_streaming(30))
    
    def test_optimal_chunk_size(self):
        """测试最佳分块大小"""
        config = PerformanceConfig()
        
        # 大文件
        large_chunk = config.get_optimal_chunk_size(300)
        self.assertGreaterEqual(large_chunk, 2000)
        
        # 小文件
        small_chunk = config.get_optimal_chunk_size(5)
        self.assertLessEqual(small_chunk, 200)
    
    def test_limits_validation(self):
        """测试限制验证"""
        config = PerformanceConfig()
        
        # 正常情况
        is_valid, errors = config.validate_limits(1000, 100, 50)
        self.assertTrue(is_valid)
        self.assertEqual(len(errors), 0)
        
        # 超限情况
        is_valid, errors = config.validate_limits(2000000, 100, 50)
        self.assertFalse(is_valid)
        self.assertGreater(len(errors), 0)


class TestHTMLConfig(unittest.TestCase):
    """测试HTML配置"""
    
    def test_default_values(self):
        """测试默认值"""
        config = HTMLConfig()
        
        self.assertEqual(config.HTML_DEFAULT_ENCODING, 'utf-8')
        self.assertTrue(config.INCLUDE_COMMENTS)
        self.assertTrue(config.INCLUDE_HYPERLINKS)
        self.assertTrue(config.RESPONSIVE_TABLES)
    
    def test_html_attributes(self):
        """测试HTML属性"""
        config = HTMLConfig()
        attributes = config.get_html_attributes()
        
        self.assertIn('border', attributes)
        self.assertIn('cellspacing', attributes)
        self.assertIn('cellpadding', attributes)
    
    def test_meta_tags(self):
        """测试元标签"""
        config = HTMLConfig()
        meta_tags = config.get_meta_tags()
        
        self.assertGreater(len(meta_tags), 0)
        self.assertTrue(any('charset' in tag for tag in meta_tags))
        self.assertTrue(any('viewport' in tag for tag in meta_tags))
    
    def test_responsive_css(self):
        """测试响应式CSS"""
        config = HTMLConfig(RESPONSIVE_TABLES=True)
        css = config.get_responsive_css()
        
        self.assertIn('@media', css)
        self.assertIn('768px', css)  # 移动断点
    
    def test_size_optimization(self):
        """测试大小优化"""
        config = HTMLConfig()
        optimized = config.optimize_for_size()
        
        self.assertFalse(optimized.INCLUDE_EMPTY_CELLS)
        self.assertFalse(optimized.INCLUDE_COMMENTS)
        self.assertTrue(optimized.MINIFY_HTML)


class TestStyleConfig(unittest.TestCase):
    """测试样式配置"""
    
    def test_default_values(self):
        """测试默认值"""
        config = StyleConfig()
        
        self.assertEqual(config.DEFAULT_FONT_FAMILY, 'Arial, sans-serif')
        self.assertEqual(config.DEFAULT_FONT_SIZE, 12)
        self.assertIsNotNone(config.BORDER_STYLE_MAPPING)
    
    def test_border_style_mapping(self):
        """测试边框样式映射"""
        config = StyleConfig()
        
        self.assertEqual(config.get_border_style('thin'), '1px solid')
        self.assertEqual(config.get_border_style('thick'), '3px solid')
        self.assertEqual(config.get_border_style('unknown'), '1px solid')
    
    def test_alignment_mapping(self):
        """测试对齐映射"""
        config = StyleConfig()
        
        self.assertEqual(config.get_alignment_style('center'), 'center')
        self.assertEqual(config.get_vertical_alignment_style('middle'), 'middle')
    
    def test_color_presets(self):
        """测试颜色预设"""
        config = StyleConfig()
        
        self.assertEqual(config.get_color('red'), '#FF0000')
        self.assertEqual(config.get_color('business_blue'), '#1f4e79')
        self.assertEqual(config.get_color('#custom'), '#custom')  # 自定义颜色
    
    def test_font_style_creation(self):
        """测试字体样式创建"""
        config = StyleConfig()
        
        styles = config.create_font_style(
            family='Helvetica',
            size=14,
            weight='bold',
            color='red'
        )
        
        self.assertEqual(styles['font-family'], 'Helvetica')
        self.assertEqual(styles['font-size'], '14px')
        self.assertEqual(styles['font-weight'], 'bold')
        self.assertEqual(styles['color'], '#FF0000')
    
    def test_css_string_conversion(self):
        """测试CSS字符串转换"""
        config = StyleConfig()
        
        styles = {'color': 'red', 'font-size': '12px'}
        css_string = config.get_css_string(styles)
        
        self.assertIn('color: red', css_string)
        self.assertIn('font-size: 12px', css_string)


class TestFormulaConfig(unittest.TestCase):
    """测试公式配置"""
    
    def test_default_values(self):
        """测试默认值"""
        config = FormulaConfig()
        
        self.assertTrue(config.ENABLE_FORMULA_PROCESSING)
        self.assertTrue(config.CALCULATE_FORMULAS)
        self.assertEqual(config.FORMULA_CACHE_SIZE, 1000)
    
    def test_supported_functions(self):
        """测试支持的函数"""
        config = FormulaConfig()
        
        self.assertTrue(config.is_function_supported('SUM'))
        self.assertTrue(config.is_function_supported('IF'))
        self.assertFalse(config.is_function_supported('UNKNOWN_FUNC'))
    
    def test_function_categories(self):
        """测试函数分类"""
        config = FormulaConfig()
        
        self.assertEqual(config.get_function_category('SUM'), 'math')
        self.assertEqual(config.get_function_category('IF'), 'logic')
        self.assertEqual(config.get_function_category('UNKNOWN'), 'unknown')
    
    def test_formula_processing_decision(self):
        """测试公式处理决策"""
        config = FormulaConfig()
        
        self.assertTrue(config.should_process_formula('=SUM(A1:A10)'))
        self.assertFalse(config.should_process_formula('not a formula'))
        
        # 测试仅支持函数模式
        config.SUPPORTED_FUNCTIONS_ONLY = True
        self.assertTrue(config.should_process_formula('=SUM(A1:A10)'))
    
    def test_performance_optimization(self):
        """测试性能优化"""
        config = FormulaConfig()
        optimized = config.optimize_for_performance()
        
        self.assertFalse(optimized.CALCULATE_FORMULAS)
        self.assertTrue(optimized.SUPPORTED_FUNCTIONS_ONLY)
        self.assertEqual(optimized.MAX_CALCULATION_DEPTH, 50)
    
    def test_accuracy_optimization(self):
        """测试准确性优化"""
        config = FormulaConfig()
        optimized = config.optimize_for_accuracy()
        
        self.assertTrue(optimized.CALCULATE_FORMULAS)
        self.assertFalse(optimized.SUPPORTED_FUNCTIONS_ONLY)
        self.assertEqual(optimized.MAX_CALCULATION_DEPTH, 200)


class TestChartConfig(unittest.TestCase):
    """测试图表配置"""
    
    def test_default_values(self):
        """测试默认值"""
        config = ChartConfig()
        
        self.assertTrue(config.ENABLE_CHART_CONVERSION)
        self.assertEqual(config.CHART_OUTPUT_FORMAT, 'svg')
        self.assertEqual(config.CHART_DEFAULT_WIDTH, 600)
        self.assertTrue(config.CHART_RESPONSIVE)
    
    def test_chart_type_support(self):
        """测试图表类型支持"""
        config = ChartConfig()
        
        self.assertTrue(config.is_chart_type_supported('pie'))
        self.assertTrue(config.is_chart_type_supported('bar'))
        self.assertFalse(config.is_chart_type_supported('unknown'))
    
    def test_color_palette(self):
        """测试配色方案"""
        config = ChartConfig()
        
        default_palette = config.get_color_palette('default')
        business_palette = config.get_color_palette('business')
        
        self.assertGreater(len(default_palette), 0)
        self.assertGreater(len(business_palette), 0)
        self.assertNotEqual(default_palette, business_palette)
    
    def test_chart_dimensions(self):
        """测试图表尺寸"""
        config = ChartConfig()
        
        # 默认尺寸
        width, height = config.get_chart_dimensions()
        self.assertEqual(width, config.CHART_DEFAULT_WIDTH)
        self.assertEqual(height, config.CHART_DEFAULT_HEIGHT)
        
        # 饼图尺寸（正方形）
        width, height = config.get_chart_dimensions('pie')
        self.assertEqual(width, height)
    
    def test_web_optimization(self):
        """测试Web优化"""
        config = ChartConfig()
        optimized = config.optimize_for_web()
        
        self.assertEqual(optimized.CHART_OUTPUT_FORMAT, 'svg')
        self.assertTrue(optimized.CHART_RESPONSIVE)
        self.assertFalse(optimized.CHART_ANIMATIONS)
    
    def test_print_optimization(self):
        """测试打印优化"""
        config = ChartConfig()
        optimized = config.optimize_for_print()
        
        self.assertEqual(optimized.CHART_OUTPUT_FORMAT, 'png')
        self.assertEqual(optimized.CHART_DPI, 300)
        self.assertFalse(optimized.CHART_RESPONSIVE)


class TestLoggingConfig(unittest.TestCase):
    """测试日志配置"""
    
    def test_default_values(self):
        """测试默认值"""
        config = LoggingConfig()
        
        self.assertEqual(config.LOG_LEVEL, 'INFO')
        self.assertTrue(config.ENABLE_CONSOLE_LOGGING)
        self.assertFalse(config.ENABLE_FILE_LOGGING)
    
    def test_log_level_conversion(self):
        """测试日志级别转换"""
        config = LoggingConfig()
        
        self.assertEqual(config.get_log_level('DEBUG'), logging.DEBUG)
        self.assertEqual(config.get_log_level('ERROR'), logging.ERROR)
    
    def test_console_handler(self):
        """测试控制台处理器"""
        config = LoggingConfig()
        handler = config.get_console_handler()
        
        self.assertIsNotNone(handler)
        self.assertIsInstance(handler, logging.StreamHandler)
    
    def test_production_optimization(self):
        """测试生产环境优化"""
        config = LoggingConfig()
        optimized = config.optimize_for_production()
        
        self.assertEqual(optimized.LOG_LEVEL, 'WARNING')
        self.assertTrue(optimized.ENABLE_FILE_LOGGING)
        self.assertFalse(optimized.ENABLE_CONSOLE_LOGGING)
    
    def test_development_optimization(self):
        """测试开发环境优化"""
        config = LoggingConfig()
        optimized = config.optimize_for_development()
        
        self.assertEqual(optimized.LOG_LEVEL, 'DEBUG')
        self.assertFalse(optimized.ENABLE_FILE_LOGGING)
        self.assertTrue(optimized.ENABLE_CONSOLE_LOGGING)


class TestUnifiedConfig(unittest.TestCase):
    """测试统一配置"""
    
    def test_initialization(self):
        """测试初始化"""
        config = UnifiedConfig()
        
        self.assertIsNotNone(config.file_format)
        self.assertIsNotNone(config.performance)
        self.assertIsNotNone(config.html)
        self.assertIsNotNone(config.style)
        self.assertIsNotNone(config.formula)
        self.assertIsNotNone(config.chart)
        self.assertIsNotNone(config.logging)
    
    def test_backward_compatibility(self):
        """测试向后兼容性"""
        config = UnifiedConfig()
        
        # 测试属性访问
        self.assertIsNotNone(config.SUPPORTED_EXCEL_FORMATS)
        self.assertEqual(config.MAX_FILE_SIZE_MB, 500)
        self.assertEqual(config.CHUNK_SIZE, 1000)
        self.assertTrue(config.ENABLE_FORMULA_PROCESSING)
    
    def test_property_setters(self):
        """测试属性设置器"""
        config = UnifiedConfig()
        
        # 测试可设置的属性
        config.CHUNK_SIZE = 2000
        self.assertEqual(config.CHUNK_SIZE, 2000)
        self.assertEqual(config.performance.CHUNK_SIZE, 2000)
    
    def test_module_config_access(self):
        """测试模块配置访问"""
        config = UnifiedConfig()
        
        # 获取特定模块配置
        perf_config = config.get_module_config('performance')
        self.assertIsInstance(perf_config, PerformanceConfig)
        
        # 更新模块配置
        new_perf_config = PerformanceConfig(CHUNK_SIZE=500)
        config.update_module_config('performance', new_perf_config)
        self.assertEqual(config.CHUNK_SIZE, 500)
    
    def test_performance_optimization(self):
        """测试性能优化"""
        config = UnifiedConfig()
        optimized = config.optimize_for_performance()
        
        self.assertIsInstance(optimized, UnifiedConfig)
        self.assertFalse(optimized.formula.CALCULATE_FORMULAS)
    
    def test_quality_optimization(self):
        """测试质量优化"""
        config = UnifiedConfig()
        optimized = config.optimize_for_quality()
        
        self.assertIsInstance(optimized, UnifiedConfig)
        self.assertTrue(optimized.formula.CALCULATE_FORMULAS)


class TestConfigFactories(unittest.TestCase):
    """测试配置工厂函数"""
    
    def test_performance_config(self):
        """测试性能配置工厂"""
        config = legacy_config.create_config_for_performance()
        self.assertIsInstance(config, legacy_config.Config)
    
    def test_quality_config(self):
        """测试质量配置工厂"""
        config = legacy_config.create_config_for_quality()
        self.assertIsInstance(config, legacy_config.Config)
    
    def test_web_config(self):
        """测试Web配置工厂"""
        config = legacy_config.create_config_for_web()
        self.assertIsInstance(config, legacy_config.Config)
    
    def test_print_config(self):
        """测试打印配置工厂"""
        config = legacy_config.create_config_for_print()
        self.assertIsInstance(config, legacy_config.Config)
    
    def test_minimal_config(self):
        """测试最小配置工厂"""
        config = legacy_config.create_minimal_config()
        self.assertIsInstance(config, legacy_config.Config)
        self.assertFalse(config.ENABLE_FORMULA_PROCESSING)
        self.assertFalse(config.ENABLE_CHART_CONVERSION)
    
    def test_development_config(self):
        """测试开发配置工厂"""
        config = legacy_config.create_development_config()
        self.assertIsInstance(config, legacy_config.Config)
    
    def test_production_config(self):
        """测试生产配置工厂"""
        config = legacy_config.create_production_config()
        self.assertIsInstance(config, legacy_config.Config)


class TestBackwardCompatibility(unittest.TestCase):
    """测试向后兼容性"""
    
    def test_themes_preserved(self):
        """测试主题保留"""
        self.assertIn('default', legacy_config.THEMES)
        self.assertIn('dark', legacy_config.THEMES)
        self.assertIn('minimal', legacy_config.THEMES)
        self.assertIn('print', legacy_config.THEMES)
    
    def test_error_messages_preserved(self):
        """测试错误信息保留"""
        self.assertIsNotNone(legacy_config.ErrorMessages.FILE_NOT_FOUND)
        self.assertIsNotNone(legacy_config.ErrorMessages.UNSUPPORTED_FORMAT)
        self.assertIsNotNone(legacy_config.ErrorMessages.TOO_MANY_ROWS)
    
    def test_config_alias(self):
        """测试Config别名"""
        config1 = legacy_config.Config()
        config2 = UnifiedConfig()
        
        # 应该是同一个类
        self.assertEqual(type(config1), type(config2))


if __name__ == '__main__':
    unittest.main() 