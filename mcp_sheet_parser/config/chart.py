# chart.py
# 图表转换相关配置

from dataclasses import dataclass
from typing import List, Dict, Tuple


@dataclass
class ChartConfig:
    """图表转换配置"""
    
    # 基本配置
    ENABLE_CHART_CONVERSION: bool = False
    CHART_OUTPUT_FORMAT: str = 'svg'  # svg, png, canvas
    CHART_EMBED_MODE: str = 'inline'  # inline, external
    
    # 尺寸配置
    CHART_DEFAULT_WIDTH: int = 600
    CHART_DEFAULT_HEIGHT: int = 400
    CHART_MIN_WIDTH: int = 200
    CHART_MIN_HEIGHT: int = 150
    CHART_MAX_WIDTH: int = 1200
    CHART_MAX_HEIGHT: int = 800
    
    # 质量配置
    CHART_QUALITY: str = 'high'  # low, medium, high
    CHART_DPI: int = 96
    CHART_ANTI_ALIASING: bool = True
    
    # 支持的图表类型
    SUPPORTED_CHART_TYPES: List[str] = None
    
    # 配色方案
    CHART_COLOR_SCHEME: str = 'default'  # default, business, modern, colorful
    COLOR_PALETTES: Dict[str, List[str]] = None
    
    # 交互配置
    CHART_RESPONSIVE: bool = True
    CHART_ANIMATIONS: bool = False
    CHART_INTERACTIVE: bool = True
    CHART_TOOLTIPS: bool = True
    
    # 导出配置
    EXPORT_CHART_DATA: bool = True
    INCLUDE_CHART_METADATA: bool = True
    
    def __post_init__(self):
        """初始化默认值"""
        if self.SUPPORTED_CHART_TYPES is None:
            self.SUPPORTED_CHART_TYPES = [
                'column', 'bar', 'line', 'pie', 'area', 'scatter',
                'bubble', 'donut', 'radar', 'gauge'
            ]
        
        if self.COLOR_PALETTES is None:
            self.COLOR_PALETTES = {
                'default': [
                    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
                    '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'
                ],
                'business': [
                    '#1f4e79', '#5b9bd5', '#70ad47', '#ffc000', '#c55a11',
                    '#264478', '#7030a0', '#0563c1', '#954f72', '#e7e6e6'
                ],
                'modern': [
                    '#667eea', '#764ba2', '#f093fb', '#f5576c', '#4facfe',
                    '#00f2fe', '#43e97b', '#38f9d7', '#ffecd2', '#fcb69f'
                ],
                'colorful': [
                    '#ff6b6b', '#4ecdc4', '#45b7d1', '#96ceb4', '#ffeaa7',
                    '#dda0dd', '#98d8c8', '#f7dc6f', '#bb8fce', '#85c1e9'
                ]
            }
    
    def is_chart_type_supported(self, chart_type: str) -> bool:
        """检查图表类型是否支持"""
        return chart_type.lower() in [t.lower() for t in self.SUPPORTED_CHART_TYPES]
    
    def get_color_palette(self, scheme: str = None) -> List[str]:
        """获取配色方案"""
        scheme = scheme or self.CHART_COLOR_SCHEME
        return self.COLOR_PALETTES.get(scheme, self.COLOR_PALETTES['default'])
    
    def get_chart_dimensions(self, chart_type: str = None) -> Tuple[int, int]:
        """获取图表尺寸"""
        width = self.CHART_DEFAULT_WIDTH
        height = self.CHART_DEFAULT_HEIGHT
        
        # 根据图表类型调整默认尺寸
        if chart_type:
            if chart_type.lower() in ['pie', 'donut', 'gauge']:
                # 圆形图表使用正方形比例
                size = min(width, height)
                width = height = size
            elif chart_type.lower() in ['bar']:
                # 条形图通常更高
                height = int(height * 1.2)
        
        # 确保在限制范围内
        width = max(self.CHART_MIN_WIDTH, min(width, self.CHART_MAX_WIDTH))
        height = max(self.CHART_MIN_HEIGHT, min(height, self.CHART_MAX_HEIGHT))
        
        return width, height
    
    def get_quality_settings(self) -> Dict[str, any]:
        """获取质量设置"""
        quality_map = {
            'low': {'dpi': 72, 'anti_aliasing': False, 'compression': 0.7},
            'medium': {'dpi': 96, 'anti_aliasing': True, 'compression': 0.8},
            'high': {'dpi': 144, 'anti_aliasing': True, 'compression': 0.9}
        }
        
        settings = quality_map.get(self.CHART_QUALITY, quality_map['medium'])
        settings['format'] = self.CHART_OUTPUT_FORMAT
        return settings
    
    def get_chart_options(self) -> Dict[str, any]:
        """获取图表选项"""
        return {
            'responsive': self.CHART_RESPONSIVE,
            'animations': self.CHART_ANIMATIONS,
            'interactive': self.CHART_INTERACTIVE,
            'tooltips': self.CHART_TOOLTIPS,
            'embed_mode': self.CHART_EMBED_MODE
        }
    
    def optimize_for_web(self) -> 'ChartConfig':
        """为网页显示优化配置"""
        optimized = ChartConfig(
            ENABLE_CHART_CONVERSION=self.ENABLE_CHART_CONVERSION,
            CHART_OUTPUT_FORMAT='svg',  # SVG适合网页
            CHART_EMBED_MODE='inline',
            CHART_DEFAULT_WIDTH=min(self.CHART_DEFAULT_WIDTH, 800),
            CHART_DEFAULT_HEIGHT=min(self.CHART_DEFAULT_HEIGHT, 600),
            CHART_MIN_WIDTH=self.CHART_MIN_WIDTH,
            CHART_MIN_HEIGHT=self.CHART_MIN_HEIGHT,
            CHART_MAX_WIDTH=self.CHART_MAX_WIDTH,
            CHART_MAX_HEIGHT=self.CHART_MAX_HEIGHT,
            CHART_QUALITY='medium',  # 平衡质量和大小
            CHART_DPI=96,
            CHART_ANTI_ALIASING=True,
            SUPPORTED_CHART_TYPES=self.SUPPORTED_CHART_TYPES,
            CHART_COLOR_SCHEME=self.CHART_COLOR_SCHEME,
            COLOR_PALETTES=self.COLOR_PALETTES,
            CHART_RESPONSIVE=True,  # 启用响应式
            CHART_ANIMATIONS=False,  # 禁用动画以提升性能
            CHART_INTERACTIVE=True,
            CHART_TOOLTIPS=True,
            EXPORT_CHART_DATA=self.EXPORT_CHART_DATA,
            INCLUDE_CHART_METADATA=self.INCLUDE_CHART_METADATA
        )
        
        return optimized
    
    def optimize_for_print(self) -> 'ChartConfig':
        """为打印优化配置"""
        optimized = ChartConfig(
            ENABLE_CHART_CONVERSION=self.ENABLE_CHART_CONVERSION,
            CHART_OUTPUT_FORMAT='png',  # PNG适合打印
            CHART_EMBED_MODE='inline',
            CHART_DEFAULT_WIDTH=self.CHART_DEFAULT_WIDTH,
            CHART_DEFAULT_HEIGHT=self.CHART_DEFAULT_HEIGHT,
            CHART_MIN_WIDTH=self.CHART_MIN_WIDTH,
            CHART_MIN_HEIGHT=self.CHART_MIN_HEIGHT,
            CHART_MAX_WIDTH=self.CHART_MAX_WIDTH,
            CHART_MAX_HEIGHT=self.CHART_MAX_HEIGHT,
            CHART_QUALITY='high',  # 高质量用于打印
            CHART_DPI=300,  # 高DPI用于打印
            CHART_ANTI_ALIASING=True,
            SUPPORTED_CHART_TYPES=self.SUPPORTED_CHART_TYPES,
            CHART_COLOR_SCHEME='business',  # 商务配色适合打印
            COLOR_PALETTES=self.COLOR_PALETTES,
            CHART_RESPONSIVE=False,  # 固定尺寸
            CHART_ANIMATIONS=False,  # 静态图表
            CHART_INTERACTIVE=False,  # 非交互式
            CHART_TOOLTIPS=False,
            EXPORT_CHART_DATA=False,  # 减少文件大小
            INCLUDE_CHART_METADATA=True
        )
        
        return optimized
    
    def get_chart_css_classes(self) -> Dict[str, str]:
        """获取图表CSS类"""
        return {
            'chart_container': 'chart-container',
            'chart_responsive': 'chart-responsive',
            'chart_legend': 'chart-legend',
            'chart_title': 'chart-title',
            'chart_tooltip': 'chart-tooltip'
        } 