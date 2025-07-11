# test_performance.py
# 性能优化模块测试

import unittest
import tempfile
import os
import time
from unittest.mock import patch, MagicMock

from mcp_sheet_parser.performance import (
    ProgressTracker, MemoryMonitor, StreamingExcelParser,
    PerformanceOptimizer, create_progress_callback,
    benchmark_performance
)
from mcp_sheet_parser.config.performance import PerformanceConfig
from mcp_sheet_parser.config import Config


class TestProgressTracker(unittest.TestCase):
    """进度跟踪器测试"""
    
    def setUp(self):
        self.tracker = ProgressTracker(100, "测试进度")
    
    def test_initialization(self):
        """测试初始化"""
        self.assertEqual(self.tracker.total_size, 100)
        self.assertEqual(self.tracker.processed, 0)
        self.assertEqual(self.tracker.description, "测试进度")
    
    def test_progress_update(self):
        """测试进度更新"""
        self.tracker.update(10)
        self.assertEqual(self.tracker.processed, 10)
        
        self.tracker.update(5)
        self.assertEqual(self.tracker.processed, 15)
    
    def test_progress_calculation(self):
        """测试进度计算"""
        self.tracker.update(25)
        progress = self.tracker._calculate_progress()
        
        self.assertEqual(progress['processed'], 25)
        self.assertEqual(progress['total'], 100)
        self.assertEqual(progress['percentage'], 25.0)
        self.assertIn('speed', progress)
        self.assertIn('eta', progress)
    
    def test_callback_registration(self):
        """测试回调注册"""
        callback_called = False
        callback_data = None
        
        def test_callback(data):
            nonlocal callback_called, callback_data
            callback_called = True
            callback_data = data
        
        self.tracker.add_callback(test_callback)
        self.tracker.update(10)
        
        # 可能由于更新间隔限制，需要等待
        time.sleep(0.2)
        self.tracker.update(1)
        
        self.assertTrue(callback_called)
        self.assertIsNotNone(callback_data)
    
    def test_completion_check(self):
        """测试完成检查"""
        self.assertFalse(self.tracker.is_complete())
        
        self.tracker.update(100)
        self.assertTrue(self.tracker.is_complete())
        
        self.tracker.update(10)  # 超过100
        self.assertTrue(self.tracker.is_complete())


class TestMemoryMonitor(unittest.TestCase):
    """内存监控器测试"""
    
    def setUp(self):
        self.monitor = MemoryMonitor(max_memory_mb=1024)
    
    def test_initialization(self):
        """测试初始化"""
        self.assertEqual(self.monitor.max_memory, 1024 * 1024 * 1024)
        self.assertEqual(self.monitor.gc_threshold, 0.8)
    
    def test_memory_usage_info(self):
        """测试内存使用信息"""
        usage = self.monitor.get_memory_usage()
        
        self.assertIn('rss', usage)
        self.assertIn('vms', usage)
        self.assertIn('rss_mb', usage)
        self.assertIn('percent', usage)
        self.assertIsInstance(usage['rss_mb'], float)
    
    def test_memory_stats(self):
        """测试内存统计"""
        stats = self.monitor.get_memory_stats()
        self.assertIsInstance(stats, str)
        self.assertIn('内存使用', stats)
        self.assertIn('MB', stats)
    
    @patch('gc.collect')
    def test_garbage_collection(self, mock_gc):
        """测试垃圾回收"""
        mock_gc.return_value = None
        
        freed_memory = self.monitor.trigger_gc()
        
        mock_gc.assert_called_once()
        self.assertIsInstance(freed_memory, float)


class TestPerformanceConfig(unittest.TestCase):
    """性能配置测试"""
    
    def setUp(self):
        self.config = PerformanceConfig()
    
    def test_default_values(self):
        """测试默认值"""
        self.assertEqual(self.config.CHUNK_SIZE, 1000)
        self.assertEqual(self.config.MAX_MEMORY_MB, 2048)
        self.assertEqual(self.config.MAX_WORKERS, 4)
        self.assertTrue(self.config.ENABLE_STREAMING)
    
    def test_file_size_adjustment(self):
        """测试文件大小调整"""
        # 大文件调整
        self.config.adjust_for_file_size(600)  # 600MB
        self.assertEqual(self.config.CHUNK_SIZE, 500)
        self.assertEqual(self.config.GC_THRESHOLD, 0.7)
        
        # 中等文件调整
        config2 = PerformanceConfig()
        config2.adjust_for_file_size(150)  # 150MB
        self.assertEqual(config2.CHUNK_SIZE, 750)
        self.assertEqual(config2.GC_THRESHOLD, 0.75)
        
        # 小文件不调整
        config3 = PerformanceConfig()
        config3.adjust_for_file_size(50)  # 50MB
        self.assertEqual(config3.CHUNK_SIZE, 1000)  # 保持默认值


class TestPerformanceOptimizer(unittest.TestCase):
    """性能优化器测试"""
    
    def setUp(self):
        self.optimizer = PerformanceOptimizer()
    
    def test_system_info(self):
        """测试系统信息获取"""
        info = self.optimizer.get_system_info()
        
        self.assertIn('cpu_count', info)
        self.assertIn('memory_total_gb', info)
        self.assertIn('memory_available_gb', info)
        self.assertIsInstance(info['cpu_count'], int)
        self.assertIsInstance(info['memory_total_gb'], float)
    
    def test_file_optimization(self):
        """测试文件优化"""
        # 创建临时文件进行测试
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            temp_file.write(b'x' * (600 * 1024 * 1024))  # 600MB 触发大文件优化
            temp_path = temp_file.name
        
        try:
            optimized_config = self.optimizer.optimize_for_file(temp_path)
            
            self.assertIsInstance(optimized_config, PerformanceConfig)
            # 大文件应该有优化配置
            self.assertEqual(optimized_config.CHUNK_SIZE, 500)  # 大文件优化后的值
            
        finally:
            os.unlink(temp_path)


class TestUtilityFunctions(unittest.TestCase):
    """工具函数测试"""
    
    def test_progress_callback_creation(self):
        """测试进度回调创建"""
        callback = create_progress_callback(verbose=False)
        self.assertTrue(callable(callback))
        
        # 测试调用不会出错
        test_data = {
            'description': '测试',
            'percentage': 50.0,
            'processed': 50,
            'total': 100,
            'speed': '10 项/秒',
            'eta': '5秒'
        }
        
        # 应该不会抛出异常
        callback(test_data)
    
    def test_benchmark_performance(self):
        """测试性能基准测试"""
        def test_function(x, y):
            time.sleep(0.01)  # 模拟一些工作
            return x + y
        
        result = benchmark_performance(test_function, 1, 2)
        
        self.assertTrue(result['success'])
        self.assertEqual(result['result'], 3)
        self.assertGreater(result['execution_time'], 0.01)
        self.assertIn('memory_delta_mb', result)
        self.assertIn('peak_memory_mb', result)
    
    def test_benchmark_with_exception(self):
        """测试异常情况下的基准测试"""
        def failing_function():
            raise ValueError("测试异常")
        
        result = benchmark_performance(failing_function)
        
        self.assertFalse(result['success'])
        self.assertIsNone(result['result'])
        self.assertIn('error', result)
        self.assertEqual(result['error'], "测试异常")


class TestStreamingParser(unittest.TestCase):
    """流式解析器测试"""
    
    def setUp(self):
        self.parser = StreamingExcelParser(chunk_size=100)
    
    def test_initialization(self):
        """测试初始化"""
        self.assertEqual(self.parser.chunk_size, 100)
        self.assertIsNotNone(self.parser.memory_monitor)
    
    def test_chunk_data_extraction(self):
        """测试数据块提取逻辑"""
        # 这里主要测试方法是否存在且可调用
        # 实际Excel文件测试需要在集成测试中进行
        self.assertTrue(hasattr(self.parser, '_extract_chunk_data'))
        self.assertTrue(hasattr(self.parser, 'parse_in_chunks'))


if __name__ == '__main__':
    unittest.main() 