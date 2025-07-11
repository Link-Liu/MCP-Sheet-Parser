# benchmark.py
# 性能基准测试模块

import argparse
from typing import Dict, Any

from .utils import setup_logger


class BenchmarkRunner:
    """基准测试运行器"""
    
    def __init__(self):
        self.logger = setup_logger(__name__)
    
    def run_benchmark(self, input_file: str) -> int:
        """
        运行性能基准测试
        
        Args:
            input_file: 输入文件路径
            
        Returns:
            int: 程序退出代码 (0=成功, 1=失败)
        """
        try:
            print("🚀 执行性能基准测试...")
            from .performance import PerformanceBenchmark
            
            benchmark = PerformanceBenchmark()
            results = benchmark.run_comprehensive_benchmark(input_file)
            
            print("\n📊 基准测试结果:")
            for metric, value in results.items():
                print(f"  {metric}: {value}")
            
            return 0
            
        except Exception as e:
            print(f"❌ 基准测试失败: {e}")
            self.logger.error(f"基准测试失败: {e}")
            return 1


# 创建全局基准测试运行器实例
benchmark_runner = BenchmarkRunner()

# 导出主要函数供main.py使用
run_benchmark = benchmark_runner.run_benchmark 