# security.py
# 简化的安全验证模块

import os
from pathlib import Path
from typing import Tuple, List


def validate_file_path(file_path: str) -> Tuple[bool, List[str]]:
    """
    验证文件路径安全性
    
    Args:
        file_path: 文件路径
        
    Returns:
        Tuple[is_safe, warnings]
    """
    warnings = []
    
    try:
        # 标准化路径
        normalized_path = os.path.normpath(os.path.abspath(file_path))
        
        # 检查路径遍历攻击
        if '..' in file_path or '//' in file_path or '\\\\' in file_path:
            warnings.append("检测到可疑的路径遍历字符")
        
        # 检查文件是否存在且可读
        if not os.path.exists(normalized_path):
            warnings.append(f"文件不存在: {file_path}")
            return False, warnings
        
        if not os.path.isfile(normalized_path):
            warnings.append(f"路径不是文件: {file_path}")
            return False, warnings
        
        if not os.access(normalized_path, os.R_OK):
            warnings.append(f"文件不可读: {file_path}")
            return False, warnings
        
        # 检查文件扩展名
        file_ext = Path(file_path).suffix.lower()
        dangerous_extensions = {
            '.exe', '.bat', '.cmd', '.com', '.scr', '.pif', '.vbs', '.js', 
            '.jar', '.app', '.deb', '.pkg', '.dmg', '.msi', '.ps1', '.sh'
        }
        
        if file_ext in dangerous_extensions:
            warnings.append(f"危险的文件扩展名: {file_ext}")
            return False, warnings
        
        # 检查文件大小 (限制100MB)
        file_size = os.path.getsize(normalized_path)
        max_size = 100 * 1024 * 1024  # 100MB
        if file_size > max_size:
            warnings.append(f"文件过大: {file_size} bytes > {max_size} bytes")
            return False, warnings
        
        # 检查文件是否为符号链接
        if os.path.islink(normalized_path):
            warnings.append("文件是符号链接，可能存在安全风险")
        
        return True, warnings
        
    except Exception as e:
        warnings.append(f"路径验证失败: {e}")
        return False, warnings


def validate_output_path(output_path: str) -> Tuple[bool, List[str]]:
    """
    验证输出路径安全性
    
    Args:
        output_path: 输出路径
        
    Returns:
        Tuple[is_safe, warnings]
    """
    warnings = []
    
    try:
        # 标准化路径
        normalized_path = os.path.normpath(os.path.abspath(output_path))
        output_dir = os.path.dirname(normalized_path)
        
        # 检查路径遍历
        if '..' in output_path:
            warnings.append("输出路径包含危险的路径遍历字符")
            return False, warnings
        
        # 检查目录是否存在，不存在则尝试创建
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, mode=0o755)
            except OSError as e:
                warnings.append(f"无法创建输出目录: {e}")
                return False, warnings
        
        # 检查写入权限
        if not os.access(output_dir, os.W_OK):
            warnings.append(f"输出目录无写入权限: {output_dir}")
            return False, warnings
        
        # 检查文件扩展名
        file_ext = Path(output_path).suffix.lower()
        allowed_output_extensions = {'.html', '.htm'}
        if file_ext not in allowed_output_extensions:
            warnings.append(f"输出文件应为HTML格式: {file_ext}")
        
        return True, warnings
        
    except Exception as e:
        warnings.append(f"输出路径验证失败: {e}")
        return False, warnings


class SecurityError(Exception):
    """安全错误异常"""
    pass 