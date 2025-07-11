# utils.py
# 基本工具函数，专注于HTML转换

import os
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional, Callable


def setup_logger(name: str, level: int = logging.INFO) -> logging.Logger:
    """设置日志记录器"""
    logger = logging.getLogger(name)
    
    if logger.handlers:
        return logger
    
    logger.setLevel(level)
    
    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)
    
    # 格式器
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)
    
    logger.addHandler(console_handler)
    
    return logger


def validate_file_path(file_path: str) -> bool:
    """
    简单的文件路径验证
    
    Args:
        file_path: 文件路径
        
    Returns:
        是否有效
    """
    try:
        return os.path.isfile(file_path) and os.access(file_path, os.R_OK)
    except Exception:
        return False


def get_file_extension(file_path: str) -> str:
    """
    获取文件扩展名
    
    Args:
        file_path: 文件路径
        
    Returns:
        小写的文件扩展名
    """
    return Path(file_path).suffix.lower()


def is_supported_format(file_path: str) -> bool:
    """
    检查文件格式是否支持
    
    Args:
        file_path: 文件路径
        
    Returns:
        是否支持
    """
    from .config import Config
    return get_file_extension(file_path) in Config.get_all_supported_formats()


def sanitize_filename(filename: str) -> str:
    """
    清理文件名
    
    Args:
        filename: 原始文件名
        
    Returns:
        清理后的文件名
    """
    import re
    
    # 移除或替换危险字符
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    
    # 移除控制字符
    filename = re.sub(r'[\x00-\x1f\x7f]', '', filename)
    
    # 限制长度
    if len(filename) > 255:
        name, ext = os.path.splitext(filename)
        max_name_len = 255 - len(ext)
        filename = name[:max_name_len] + ext
    
    return filename.strip()


def ensure_output_dir(output_path: str) -> None:
    """
    确保输出目录存在
    
    Args:
        output_path: 输出文件路径
    """
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)


def format_file_size(size_bytes: int) -> str:
    """
    格式化文件大小
    
    Args:
        size_bytes: 字节数
        
    Returns:
        格式化的大小字符串
    """
    if size_bytes == 0:
        return "0B"
    
    size_names = ["B", "KB", "MB", "GB"]
    i = 0
    
    size = float(size_bytes)
    while size >= 1024 and i < len(size_names) - 1:
        size /= 1024.0
        i += 1
    
    # 对于字节，显示整数
    if i == 0:
        return f"{int(size)}{size_names[i]}"
    else:
        return f"{size:.1f}{size_names[i]}"


def get_file_info(file_path: str) -> Dict[str, Any]:
    """
    获取文件基本信息
    
    Args:
        file_path: 文件路径
        
    Returns:
        文件信息字典
    """
    if not os.path.exists(file_path):
        return {"error": "文件不存在"}
    
    try:
        stat = os.stat(file_path)
        
        return {
            "path": file_path,
            "name": os.path.basename(file_path),
            "size": stat.st_size,
            "size_formatted": format_file_size(stat.st_size),
            "extension": get_file_extension(file_path),
            "is_supported": is_supported_format(file_path),
            "modified_time": stat.st_mtime
        }
    except Exception as e:
        return {"error": f"获取文件信息失败: {e}"}


def batch_process_files(file_paths: List[str], 
                       process_func: Callable[[str], Any],
                       progress_callback: Optional[Callable] = None) -> List[Dict[str, Any]]:
    """
    批量处理文件
    
    Args:
        file_paths: 文件路径列表
        process_func: 处理函数
        progress_callback: 进度回调函数
        
    Returns:
        处理结果列表
    """
    results = []
    total = len(file_paths)
    
    for i, file_path in enumerate(file_paths):
        try:
            result = process_func(file_path)
            results.append({
                "file_path": file_path,
                "success": True,
                "result": result
            })
        except Exception as e:
            results.append({
                "file_path": file_path,
                "success": False,
                "error": str(e)
            })
        
        # 调用进度回调
        if progress_callback:
            progress_callback(i + 1, total, file_path)
    
    return results
