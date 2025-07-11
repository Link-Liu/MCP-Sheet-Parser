# performance.py
# 性能优化模块 - 大文件处理、进度跟踪、内存优化

import gc
import time
import psutil
import threading
import logging
from typing import Iterator, Callable, Optional, Dict, Any, List, Tuple
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl
from openpyxl.utils import get_column_letter


class ProgressTracker:
    """进度跟踪器"""
    
    def __init__(self, total_size: int, description: str = "处理中"):
        self.total_size = total_size
        self.processed = 0
        self.start_time = time.time()
        self.description = description
        self.callbacks: List[Callable] = []
        self.last_update_time = 0
        self.update_interval = 0.1  # 100ms更新间隔
        
    def add_callback(self, callback: Callable[[Dict[str, Any]], None]):
        """添加进度回调函数"""
        self.callbacks.append(callback)
    
    def update(self, increment: int = 1):
        """更新进度"""
        self.processed += increment
        current_time = time.time()
        
        # 限制更新频率
        if current_time - self.last_update_time < self.update_interval:
            return
        
        self.last_update_time = current_time
        progress_info = self._calculate_progress()
        
        # 调用所有回调函数
        for callback in self.callbacks:
            try:
                callback(progress_info)
            except Exception as e:
                logging.warning(f"进度回调函数执行失败: {e}")
    
    def _calculate_progress(self) -> Dict[str, Any]:
        """计算进度信息"""
        elapsed_time = time.time() - self.start_time
        progress_ratio = min(self.processed / self.total_size, 1.0) if self.total_size > 0 else 0
        percentage = progress_ratio * 100
        
        # 计算ETA
        if progress_ratio > 0 and progress_ratio < 1:
            eta_seconds = (elapsed_time / progress_ratio) * (1 - progress_ratio)
            eta = self._format_time(eta_seconds)
        else:
            eta = "计算中..." if progress_ratio == 0 else "完成"
        
        # 计算处理速度
        speed = self.processed / elapsed_time if elapsed_time > 0 else 0
        
        return {
            'description': self.description,
            'processed': self.processed,
            'total': self.total_size,
            'percentage': percentage,
            'elapsed_time': self._format_time(elapsed_time),
            'eta': eta,
            'speed': f"{speed:.1f} 项/秒" if speed > 0 else "0 项/秒"
        }
    
    def _format_time(self, seconds: float) -> str:
        """格式化时间显示"""
        if seconds < 60:
            return f"{seconds:.1f}秒"
        elif seconds < 3600:
            return f"{seconds/60:.1f}分钟"
        else:
            return f"{seconds/3600:.1f}小时"
    
    def is_complete(self) -> bool:
        """检查是否完成"""
        return self.processed >= self.total_size


class MemoryMonitor:
    """内存监控器"""
    
    def __init__(self, max_memory_mb: int = 2048, gc_threshold: float = 0.8):
        self.max_memory = max_memory_mb * 1024 * 1024  # 转换为字节
        self.gc_threshold = gc_threshold
        self.process = psutil.Process()
        self.logger = logging.getLogger(__name__)
        
    def get_memory_usage(self) -> Dict[str, Any]:
        """获取当前内存使用情况"""
        memory_info = self.process.memory_info()
        memory_percent = self.process.memory_percent()
        
        return {
            'rss': memory_info.rss,  # 物理内存
            'vms': memory_info.vms,  # 虚拟内存
            'rss_mb': memory_info.rss / (1024 * 1024),
            'vms_mb': memory_info.vms / (1024 * 1024),
            'percent': memory_percent,
            'max_mb': self.max_memory / (1024 * 1024)
        }
    
    def check_memory_limit(self) -> bool:
        """检查是否超过内存限制"""
        current_memory = self.process.memory_info().rss
        return current_memory > self.max_memory
    
    def should_trigger_gc(self) -> bool:
        """判断是否应该触发垃圾回收"""
        current_memory = self.process.memory_info().rss
        threshold = self.max_memory * self.gc_threshold
        return current_memory > threshold
    
    def trigger_gc(self):
        """触发垃圾回收"""
        before_memory = self.process.memory_info().rss
        gc.collect()
        after_memory = self.process.memory_info().rss
        freed_mb = (before_memory - after_memory) / (1024 * 1024)
        
        self.logger.info(f"垃圾回收完成，释放内存: {freed_mb:.2f}MB")
        return freed_mb
    
    def get_memory_stats(self) -> str:
        """获取内存统计信息"""
        usage = self.get_memory_usage()
        return (f"内存使用: {usage['rss_mb']:.1f}MB / "
                f"{usage['max_mb']:.1f}MB ({usage['percent']:.1f}%)")


class StreamingExcelParser:
    """流式Excel解析器"""
    
    def __init__(self, chunk_size: int = 1000, enable_memory_monitor: bool = True):
        self.chunk_size = chunk_size
        self.memory_monitor = MemoryMonitor() if enable_memory_monitor else None
        self.logger = logging.getLogger(__name__)
        
    def parse_in_chunks(self, file_path: str, 
                       progress_callback: Optional[Callable] = None) -> Iterator[Dict[str, Any]]:
        """分块解析Excel文件"""
        try:
            # 使用只读模式打开文件以节省内存
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            
            for ws_index, ws in enumerate(wb.worksheets):
                yield from self._parse_worksheet_chunks(ws, ws_index, progress_callback)
                
                # 内存检查和垃圾回收
                if self.memory_monitor and self.memory_monitor.should_trigger_gc():
                    self.memory_monitor.trigger_gc()
            
            wb.close()  # 确保关闭文件
            
        except Exception as e:
            self.logger.error(f"流式解析文件失败 {file_path}: {e}")
            raise
    
    def _parse_worksheet_chunks(self, worksheet, ws_index: int,
                              progress_callback: Optional[Callable] = None) -> Iterator[Dict[str, Any]]:
        """分块解析单个工作表"""
        total_rows = worksheet.max_row
        total_chunks = (total_rows + self.chunk_size - 1) // self.chunk_size
        
        # 创建进度跟踪器
        if progress_callback:
            tracker = ProgressTracker(total_rows, f"解析工作表 {worksheet.title}")
            tracker.add_callback(progress_callback)
        else:
            tracker = None
        
        for chunk_index in range(total_chunks):
            start_row = chunk_index * self.chunk_size + 1
            end_row = min(start_row + self.chunk_size - 1, total_rows)
            
            chunk_data = self._extract_chunk_data(worksheet, start_row, end_row)
            
            yield {
                'worksheet_index': ws_index,
                'worksheet_name': worksheet.title,
                'chunk_index': chunk_index,
                'start_row': start_row,
                'end_row': end_row,
                'total_chunks': total_chunks,
                'data': chunk_data['data'],
                'styles': chunk_data['styles'],
                'comments': chunk_data['comments'],
                'hyperlinks': chunk_data['hyperlinks'],
                'merged_cells': chunk_data['merged_cells']
            }
            
            # 更新进度
            if tracker:
                tracker.update(end_row - start_row + 1)
            
            # 内存检查
            if self.memory_monitor and self.memory_monitor.check_memory_limit():
                self.logger.warning("内存使用超限，强制垃圾回收")
                self.memory_monitor.trigger_gc()
    
    def _extract_chunk_data(self, worksheet, start_row: int, end_row: int) -> Dict[str, Any]:
        """提取数据块"""
        from .parser import _get_cell_style, clean_cell_value
        
        data = []
        styles = []
        comments = {}
        hyperlinks = {}
        merged_cells = []
        
        # 处理数据行
        for row_idx in range(start_row, end_row + 1):
            data_row = []
            style_row = []
            
            for col_idx in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                
                # 提取数据
                cell_value = cell.value if cell.value is not None else ''
                data_row.append(clean_cell_value(cell_value))
                
                # 提取样式
                style = _get_cell_style(cell)
                style_row.append(style)
                
                # 提取注释和超链接
                cell_key = f"{row_idx-1}_{col_idx-1}"
                
                if 'comment' in style and style['comment']:
                    comments[cell_key] = style['comment']
                
                if 'hyperlink' in style and style['hyperlink']:
                    hyperlinks[cell_key] = style['hyperlink']
            
            data.append(data_row)
            styles.append(style_row)
        
        # 处理合并单元格（在当前块范围内的）
        for merged in worksheet.merged_cells.ranges:
            if (start_row <= merged.min_row <= end_row or 
                start_row <= merged.max_row <= end_row):
                merged_cells.append((
                    merged.min_row - 1, merged.min_col - 1,
                    merged.max_row - 1, merged.max_col - 1
                ))
        
        return {
            'data': data,
            'styles': styles,
            'comments': comments,
            'hyperlinks': hyperlinks,
            'merged_cells': merged_cells
        }


class ParallelSheetProcessor:
    """并行工作表处理器"""
    
    def __init__(self, max_workers: int = 4):
        self.max_workers = max_workers
        self.logger = logging.getLogger(__name__)
    
    def process_sheets_parallel(self, file_path: str, 
                               process_func: Callable,
                               progress_callback: Optional[Callable] = None) -> List[Any]:
        """并行处理多个工作表"""
        try:
            # 首先获取所有工作表信息
            wb = openpyxl.load_workbook(file_path, read_only=True)
            sheet_names = [ws.title for ws in wb.worksheets]
            wb.close()
            
            if not sheet_names:
                return []
            
            # 创建进度跟踪器
            if progress_callback:
                tracker = ProgressTracker(len(sheet_names), "并行处理工作表")
                tracker.add_callback(progress_callback)
            else:
                tracker = None
            
            results = []
            
            # 使用线程池并行处理
            with ThreadPoolExecutor(max_workers=min(self.max_workers, len(sheet_names))) as executor:
                # 提交所有任务
                future_to_sheet = {
                    executor.submit(process_func, file_path, i): (i, name)
                    for i, name in enumerate(sheet_names)
                }
                
                # 收集结果
                for future in as_completed(future_to_sheet):
                    sheet_index, sheet_name = future_to_sheet[future]
                    
                    try:
                        result = future.result()
                        results.append((sheet_index, result))
                        
                        if tracker:
                            tracker.update()
                        
                        self.logger.info(f"工作表 {sheet_name} 处理完成")
                        
                    except Exception as e:
                        self.logger.error(f"工作表 {sheet_name} 处理失败: {e}")
                        results.append((sheet_index, None))
            
            # 按原始顺序排序结果
            results.sort(key=lambda x: x[0])
            return [result for _, result in results if result is not None]
            
        except Exception as e:
            self.logger.error(f"并行处理失败: {e}")
            raise


class PerformanceConfig:
    """性能配置类"""
    
    def __init__(self):
        # 分块处理配置
        self.CHUNK_SIZE = 1000              # 每块处理的行数
        self.ENABLE_STREAMING = True        # 启用流式处理
        
        # 内存管理配置
        self.MAX_MEMORY_MB = 2048          # 最大内存使用（MB）
        self.GC_THRESHOLD = 0.8            # 垃圾回收触发阈值
        self.ENABLE_MEMORY_MONITOR = True   # 启用内存监控
        
        # 并行处理配置
        self.MAX_WORKERS = 4               # 最大并行工作线程数
        self.ENABLE_PARALLEL = True        # 启用并行处理
        
        # 进度跟踪配置
        self.PROGRESS_UPDATE_INTERVAL = 0.1  # 进度更新间隔（秒）
        self.ENABLE_PROGRESS_TRACKING = True # 启用进度跟踪
        
        # 缓存配置
        self.ENABLE_CACHING = True         # 启用缓存
        self.CACHE_SIZE_MB = 256          # 缓存大小（MB）
        
    def adjust_for_file_size(self, file_size_mb: float):
        """根据文件大小调整配置"""
        if file_size_mb > 500:  # 大文件
            self.CHUNK_SIZE = 500
            self.MAX_MEMORY_MB = min(4096, self.MAX_MEMORY_MB * 2)
            self.GC_THRESHOLD = 0.7
        elif file_size_mb > 100:  # 中等文件
            self.CHUNK_SIZE = 750
            self.GC_THRESHOLD = 0.75
        # 小文件使用默认配置


def create_progress_callback(verbose: bool = True) -> Callable:
    """创建标准进度回调函数"""
    
    def callback(progress_info: Dict[str, Any]):
        if verbose:
            print(f"\r{progress_info['description']}: "
                  f"{progress_info['percentage']:.1f}% "
                  f"({progress_info['processed']}/{progress_info['total']}) "
                  f"- {progress_info['speed']} "
                  f"- 剩余: {progress_info['eta']}", end='')
            
            if progress_info['percentage'] >= 100:
                print()  # 完成时换行
    
    return callback


def benchmark_performance(func: Callable, *args, **kwargs) -> Dict[str, Any]:
    """性能基准测试"""
    memory_monitor = MemoryMonitor()
    
    # 记录开始状态
    start_time = time.time()
    start_memory = memory_monitor.get_memory_usage()
    
    try:
        # 执行函数
        result = func(*args, **kwargs)
        
        # 记录结束状态
        end_time = time.time()
        end_memory = memory_monitor.get_memory_usage()
        
        return {
            'result': result,
            'execution_time': end_time - start_time,
            'memory_delta_mb': end_memory['rss_mb'] - start_memory['rss_mb'],
            'peak_memory_mb': end_memory['rss_mb'],
            'success': True
        }
        
    except Exception as e:
        end_time = time.time()
        return {
            'result': None,
            'execution_time': end_time - start_time,
            'error': str(e),
            'success': False
        }


class PerformanceOptimizer:
    """性能优化器主类"""
    
    def __init__(self, config: Optional[PerformanceConfig] = None):
        self.config = config or PerformanceConfig()
        self.logger = logging.getLogger(__name__)
        
    def optimize_for_file(self, file_path: str) -> PerformanceConfig:
        """为特定文件优化配置"""
        file_size_mb = Path(file_path).stat().st_size / (1024 * 1024)
        optimized_config = PerformanceConfig()
        optimized_config.adjust_for_file_size(file_size_mb)
        
        self.logger.info(f"为文件 {file_path} ({file_size_mb:.1f}MB) 优化配置")
        self.logger.info(f"块大小: {optimized_config.CHUNK_SIZE}, "
                        f"最大内存: {optimized_config.MAX_MEMORY_MB}MB")
        
        return optimized_config
    
    def get_system_info(self) -> Dict[str, Any]:
        """获取系统信息"""
        return {
            'cpu_count': psutil.cpu_count(),
            'cpu_percent': psutil.cpu_percent(),
            'memory_total_gb': psutil.virtual_memory().total / (1024**3),
            'memory_available_gb': psutil.virtual_memory().available / (1024**3),
            'memory_percent': psutil.virtual_memory().percent
        } 


class PerformanceBenchmark:
    """性能基准测试类"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def run_comprehensive_benchmark(self, file_path: str) -> Dict[str, Any]:
        """运行综合性能基准测试"""
        results = {}
        
        try:
            # 文件信息
            file_size_mb = Path(file_path).stat().st_size / (1024 * 1024)
            results['文件大小(MB)'] = f"{file_size_mb:.2f}"
            
            # 测试标准解析
            results.update(self._benchmark_standard_parsing(file_path))
            
            # 测试流式解析（如果文件较大）
            if file_size_mb > 50:
                results.update(self._benchmark_streaming_parsing(file_path))
            
            # 系统信息
            system_info = self._get_system_benchmark()
            results.update(system_info)
            
        except Exception as e:
            results['错误'] = str(e)
            
        return results
    
    def _benchmark_standard_parsing(self, file_path: str) -> Dict[str, Any]:
        """基准测试标准解析"""
        from .parser import SheetParser
        from .config import Config
        
        config = Config()
        config.ENABLE_PERFORMANCE_MODE = False
        
        def parse_standard():
            parser = SheetParser(file_path, config)
            return parser.parse()
        
        result = benchmark_performance(parse_standard)
        
        return {
            '标准解析时间(秒)': f"{result['execution_time']:.2f}",
            '标准解析内存(MB)': f"{result['memory_delta_mb']:.2f}",
            '标准解析成功': result['success']
        }
    
    def _benchmark_streaming_parsing(self, file_path: str) -> Dict[str, Any]:
        """基准测试流式解析"""
        from .parser import SheetParser
        from .config import Config
        
        config = Config()
        config.ENABLE_PERFORMANCE_MODE = True
        config.CHUNK_SIZE = 1000
        
        def parse_streaming():
            parser = SheetParser(file_path, config)
            return parser.parse()
        
        result = benchmark_performance(parse_streaming)
        
        return {
            '流式解析时间(秒)': f"{result['execution_time']:.2f}",
            '流式解析内存(MB)': f"{result['memory_delta_mb']:.2f}",
            '流式解析成功': result['success']
        }
    
    def _get_system_benchmark(self) -> Dict[str, Any]:
        """获取系统基准信息"""
        optimizer = PerformanceOptimizer()
        system_info = optimizer.get_system_info()
        
        return {
            'CPU核心数': system_info['cpu_count'],
            'CPU使用率(%)': f"{system_info['cpu_percent']:.1f}",
            '总内存(GB)': f"{system_info['memory_total_gb']:.1f}",
            '可用内存(GB)': f"{system_info['memory_available_gb']:.1f}",
            '内存使用率(%)': f"{system_info['memory_percent']:.1f}"
        } 