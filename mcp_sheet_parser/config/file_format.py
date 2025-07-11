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