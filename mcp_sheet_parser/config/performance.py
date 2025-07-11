# performance.py
# 性能相关配置

from dataclasses import dataclass
from typing import Optional


@dataclass
class PerformanceConfig:
    """性能配置"""
    
    # 文件大小限制
    MAX_FILE_SIZE_MB: int = 500
    MAX_ROWS: int = 1000000
    MAX_COLS: int = 16384
    
    # 性能优化配置
    ENABLE_PERFORMANCE_MODE: bool = True
    CHUNK_SIZE: int = 1000
    MAX_MEMORY_MB: int = 2048
    MAX_WORKERS: int = 4  # 添加MAX_WORKERS属性
    ENABLE_PROGRESS_TRACKING: bool = True
    ENABLE_PARALLEL_PROCESSING: bool = True
    
    # 缓存配置
    ENABLE_CACHING: bool = True
    CACHE_SIZE_MB: int = 256
    CACHE_CLEANUP_INTERVAL: int = 300  # 秒
    
    # 垃圾回收配置
    GC_THRESHOLD: float = 0.8
    ENABLE_STREAMING: bool = True
    
    def is_large_file(self, size_mb: float) -> bool:
        """判断是否为大文件"""
        return size_mb > 100
    
    def should_use_streaming(self, size_mb: float) -> bool:
        """判断是否应该使用流式处理"""
        return size_mb > 50 and self.ENABLE_PERFORMANCE_MODE
    
    def get_optimal_chunk_size(self, file_size_mb: float, mode: str = 'auto') -> int:
        """获取最佳分块大小"""
        if mode == 'fast':
            return min(self.CHUNK_SIZE, 500)
        elif mode == 'memory':
            return max(self.CHUNK_SIZE, 2000)
        else:  # auto
            if file_size_mb > 200:
                return max(self.CHUNK_SIZE, 2000)  # 大文件用大块
            elif file_size_mb < 10:
                return min(self.CHUNK_SIZE, 200)   # 小文件用小块
            else:
                return self.CHUNK_SIZE
    
    def validate_limits(self, rows: int, cols: int, file_size_mb: float) -> tuple[bool, list[str]]:
        """验证是否超过性能限制"""
        errors = []
        
        if rows > self.MAX_ROWS:
            errors.append(f"行数超过限制: {rows} > {self.MAX_ROWS}")
        
        if cols > self.MAX_COLS:
            errors.append(f"列数超过限制: {cols} > {self.MAX_COLS}")
        
        if file_size_mb > self.MAX_FILE_SIZE_MB:
            errors.append(f"文件大小超过限制: {file_size_mb}MB > {self.MAX_FILE_SIZE_MB}MB")
        
        return len(errors) == 0, errors
    
    def get_memory_config(self) -> dict:
        """获取内存配置"""
        return {
            'max_memory_mb': self.MAX_MEMORY_MB,
            'chunk_size': self.CHUNK_SIZE,
            'enable_caching': self.ENABLE_CACHING,
            'cache_size_mb': self.CACHE_SIZE_MB
        }
    
    def optimize_for_file_size(self, file_size_mb: float) -> 'PerformanceConfig':
        """根据文件大小优化配置"""
        # 创建配置副本
        optimized = PerformanceConfig(
            MAX_FILE_SIZE_MB=self.MAX_FILE_SIZE_MB,
            MAX_ROWS=self.MAX_ROWS,
            MAX_COLS=self.MAX_COLS,
            ENABLE_PERFORMANCE_MODE=self.ENABLE_PERFORMANCE_MODE,
            CHUNK_SIZE=self.get_optimal_chunk_size(file_size_mb),
            MAX_MEMORY_MB=self.MAX_MEMORY_MB,
            ENABLE_PROGRESS_TRACKING=self.ENABLE_PROGRESS_TRACKING,
            ENABLE_PARALLEL_PROCESSING=self.ENABLE_PARALLEL_PROCESSING,
            ENABLE_CACHING=self.ENABLE_CACHING,
            CACHE_SIZE_MB=self.CACHE_SIZE_MB,
            CACHE_CLEANUP_INTERVAL=self.CACHE_CLEANUP_INTERVAL
        )
        
        # 大文件优化
        if file_size_mb > 100:
            optimized.ENABLE_PARALLEL_PROCESSING = True
            optimized.ENABLE_CACHING = True
            optimized.CACHE_SIZE_MB = min(512, self.CACHE_SIZE_MB * 2)
        
        return optimized 
    
    def adjust_for_file_size(self, file_size_mb: float):
        """根据文件大小调整配置（就地修改）"""
        if file_size_mb > 500:  # 大文件
            self.CHUNK_SIZE = 500
            self.GC_THRESHOLD = 0.7
            self.ENABLE_STREAMING = True
        elif file_size_mb > 100:  # 中等文件
            self.CHUNK_SIZE = 750
            self.GC_THRESHOLD = 0.75
        # 小文件保持默认值 