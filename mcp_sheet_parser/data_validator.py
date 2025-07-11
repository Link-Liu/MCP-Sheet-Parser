# data_validator.py
# 简化的数据验证模块，专注于HTML安全

import re
import unicodedata
from typing import Any


def clean_cell_value(value: Any) -> str:
    """
    清理单元格值用于HTML输出
    
    Args:
        value: 原始值
        
    Returns:
        清理后的字符串
    """
    if value is None:
        return ''
    
    str_value = str(value).strip()
    
    if not str_value:
        return ''
    
    # 移除控制字符
    cleaned = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', str_value)
    
    # 标准化Unicode字符
    cleaned = unicodedata.normalize('NFKC', cleaned)
    
    # 移除多余的空白字符
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    
    return cleaned


def escape_html(value: Any) -> str:
    """
    HTML实体转义
    
    Args:
        value: 原始值
        
    Returns:
        转义后的字符串
    """
    if value is None:
        return ''
    
    str_value = str(value)
    
    # HTML实体转义
    html_escape_table = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#x27;',
        '/': '&#x2F;'
    }
    
    # 替换HTML特殊字符
    for char, escape in html_escape_table.items():
        str_value = str_value.replace(char, escape)
    
    # 移除危险字符
    str_value = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', str_value)
    
    # 限制长度
    max_length = 10000
    if len(str_value) > max_length:
        str_value = str_value[:max_length] + '...'
    
    return str_value


def sanitize_for_html(value: Any) -> str:
    """
    为HTML输出清理和安全化数据
    
    Args:
        value: 原始值
        
    Returns:
        安全的HTML字符串
    """
    cleaned = clean_cell_value(value)
    return escape_html(cleaned) 