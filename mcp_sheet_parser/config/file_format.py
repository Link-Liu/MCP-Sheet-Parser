# file_format.py
# 文件格式相关配置

from typing import Set
from dataclasses import dataclass


@dataclass
class FileFormatConfig:
    """文件格式配置"""
    
    # 支持的文件格式
    SUPPORTED_EXCEL_FORMATS: Set[str] = None
    SUPPORTED_CSV_FORMATS: Set[str] = None
    SUPPORTED_WPS_FORMATS: Set[str] = None
    
    # 格式支持级别
    FORMAT_SUPPORT_LEVELS: dict = None
    
    def __post_init__(self):
        """初始化默认值"""
        if self.SUPPORTED_EXCEL_FORMATS is None:
            self.SUPPORTED_EXCEL_FORMATS = {
                '.xlsx', '.xls', '.xlt', '.xlsb', '.xlsm', '.xltm', '.xltx'
            }
        
        if self.SUPPORTED_CSV_FORMATS is None:
            self.SUPPORTED_CSV_FORMATS = {
                '.csv'
            }
        
        if self.SUPPORTED_WPS_FORMATS is None:
            self.SUPPORTED_WPS_FORMATS = {
                '.et', '.ett', '.ets'
            }
        
        if self.FORMAT_SUPPORT_LEVELS is None:
            self.FORMAT_SUPPORT_LEVELS = {
                # Excel格式支持级别
                '.xlsx': 'FULL',      # 完全支持
                '.xlsm': 'FULL',      # 完全支持
                '.xltx': 'FULL',      # 完全支持
                '.xltm': 'FULL',      # 完全支持
                '.xls': 'BASIC',      # 基础支持
                '.xlt': 'BASIC',      # 基础支持
                '.xlsb': 'BASIC',     # 基础支持（已增强）
                
                # CSV格式支持级别
                '.csv': 'FULL',       # 完全支持
                
                # WPS格式支持级别（已提升）
                '.et': 'BASIC',       # 基础支持
                '.ett': 'BASIC',      # 基础支持
                '.ets': 'BASIC'       # 基础支持
            }
    
    def get_all_supported_formats(self) -> Set[str]:
        """获取所有支持的文件格式"""
        return (self.SUPPORTED_EXCEL_FORMATS | 
                self.SUPPORTED_CSV_FORMATS | 
                self.SUPPORTED_WPS_FORMATS)
    
    def is_excel_format(self, extension: str) -> bool:
        """检查是否为Excel格式"""
        return extension.lower() in self.SUPPORTED_EXCEL_FORMATS
    
    def is_csv_format(self, extension: str) -> bool:
        """检查是否为CSV格式"""
        return extension.lower() in self.SUPPORTED_CSV_FORMATS
    
    def is_wps_format(self, extension: str) -> bool:
        """检查是否为WPS格式"""
        return extension.lower() in self.SUPPORTED_WPS_FORMATS
    
    def is_supported_format(self, extension: str) -> bool:
        """检查是否为支持的格式"""
        return extension.lower() in self.get_all_supported_formats()
    
    def get_format_type(self, extension: str) -> str:
        """获取文件格式类型"""
        ext = extension.lower()
        if ext in self.SUPPORTED_EXCEL_FORMATS:
            return 'excel'
        elif ext in self.SUPPORTED_CSV_FORMATS:
            return 'csv'
        elif ext in self.SUPPORTED_WPS_FORMATS:
            return 'wps'
        else:
            return 'unknown'
    
    def get_support_level(self, extension: str) -> str:
        """获取格式支持级别"""
        ext = extension.lower()
        return self.FORMAT_SUPPORT_LEVELS.get(ext, 'UNKNOWN')
    
    def is_fully_supported(self, extension: str) -> bool:
        """检查是否为完全支持的格式"""
        return self.get_support_level(extension) == 'FULL'
    
    def is_basic_supported(self, extension: str) -> bool:
        """检查是否为基础支持的格式"""
        return self.get_support_level(extension) == 'BASIC'
    
    def get_support_badge(self, extension: str) -> str:
        """获取格式支持徽章文本"""
        level = self.get_support_level(extension)
        if level == 'FULL':
            return '✅ 完全支持'
        elif level == 'BASIC':
            return '⚠️ 基础支持'
        else:
            return '❌ 不支持' 