# style.py
# 样式相关配置

from dataclasses import dataclass
from typing import Dict


@dataclass
class StyleConfig:
    """样式配置"""
    
    # 默认样式
    DEFAULT_FONT_FAMILY: str = 'Arial, sans-serif'
    DEFAULT_FONT_SIZE: int = 12
    DEFAULT_BORDER_COLOR: str = '#000000'
    DEFAULT_BACKGROUND_COLOR: str = '#FFFFFF'
    DEFAULT_TEXT_COLOR: str = '#000000'
    
    # 边框样式映射
    BORDER_STYLE_MAPPING: Dict[str, str] = None
    
    # 对齐方式映射
    ALIGNMENT_MAPPING: Dict[str, str] = None
    VERTICAL_ALIGNMENT_MAPPING: Dict[str, str] = None
    
    # 颜色预设
    COLOR_PRESETS: Dict[str, str] = None
    
    def __post_init__(self):
        """初始化默认值"""
        if self.BORDER_STYLE_MAPPING is None:
            self.BORDER_STYLE_MAPPING = {
                'thin': '1px solid',
                'medium': '2px solid',
                'thick': '3px solid',
                'dashed': '1px dashed',
                'dotted': '1px dotted',
                'double': '3px double',
                'hair': '1px solid',
                'mediumDashed': '2px dashed',
                'dashDot': '1px dashed',
                'mediumDashDot': '2px dashed',
                'dashDotDot': '1px dashed',
                'mediumDashDotDot': '2px dashed',
                'slantDashDot': '1px dashed'
            }
        
        if self.ALIGNMENT_MAPPING is None:
            self.ALIGNMENT_MAPPING = {
                'left': 'left',
                'center': 'center',
                'right': 'right',
                'justify': 'justify',
                'centerContinuous': 'center',
                'distributed': 'justify',
                'fill': 'left',
                'general': 'left'
            }
        
        if self.VERTICAL_ALIGNMENT_MAPPING is None:
            self.VERTICAL_ALIGNMENT_MAPPING = {
                'top': 'top',
                'center': 'middle',
                'middle': 'middle',  # 添加missing映射
                'bottom': 'bottom',
                'justify': 'middle',
                'distributed': 'middle'
            }
        
        if self.COLOR_PRESETS is None:
            self.COLOR_PRESETS = {
                # 基础颜色
                'black': '#000000',
                'white': '#FFFFFF',
                'red': '#FF0000',
                'green': '#008000',
                'blue': '#0000FF',
                'yellow': '#FFFF00',
                'cyan': '#00FFFF',
                'magenta': '#FF00FF',
                
                # 商务颜色
                'business_blue': '#1f4e79',
                'business_gray': '#595959',
                'business_light_blue': '#5b9bd5',
                'business_green': '#70ad47',
                'business_orange': '#ffc000',
                
                # 警告颜色
                'success': '#28a745',
                'warning': '#ffc107',
                'danger': '#dc3545',
                'info': '#17a2b8'
            }
    
    def get_border_style(self, style_name: str) -> str:
        """获取边框样式"""
        return self.BORDER_STYLE_MAPPING.get(style_name, '1px solid')
    
    def get_alignment_style(self, alignment: str) -> str:
        """获取对齐样式"""
        return self.ALIGNMENT_MAPPING.get(alignment, 'left')
    
    def get_vertical_alignment_style(self, alignment: str) -> str:
        """获取垂直对齐样式"""
        return self.VERTICAL_ALIGNMENT_MAPPING.get(alignment, 'top')
    
    def get_color(self, color_name: str) -> str:
        """获取预设颜色"""
        return self.COLOR_PRESETS.get(color_name, color_name)
    
    def create_font_style(self, 
                         family: str = None, 
                         size: int = None, 
                         weight: str = 'normal',
                         style: str = 'normal',
                         color: str = None) -> Dict[str, str]:
        """创建字体样式"""
        styles = {}
        
        if family:
            styles['font-family'] = family
        else:
            styles['font-family'] = self.DEFAULT_FONT_FAMILY
        
        if size:
            styles['font-size'] = f'{size}px'
        else:
            styles['font-size'] = f'{self.DEFAULT_FONT_SIZE}px'
        
        if weight != 'normal':
            styles['font-weight'] = weight
        
        if style != 'normal':
            styles['font-style'] = style
        
        if color:
            styles['color'] = self.get_color(color)
        else:
            styles['color'] = self.DEFAULT_TEXT_COLOR
        
        return styles
    
    def create_border_style(self, 
                           top: str = None,
                           right: str = None,
                           bottom: str = None,
                           left: str = None,
                           color: str = None) -> Dict[str, str]:
        """创建边框样式"""
        styles = {}
        border_color = color or self.DEFAULT_BORDER_COLOR
        
        if top:
            styles['border-top'] = f'{self.get_border_style(top)} {border_color}'
        if right:
            styles['border-right'] = f'{self.get_border_style(right)} {border_color}'
        if bottom:
            styles['border-bottom'] = f'{self.get_border_style(bottom)} {border_color}'
        if left:
            styles['border-left'] = f'{self.get_border_style(left)} {border_color}'
        
        return styles
    
    def create_cell_style(self,
                         font_family: str = None,
                         font_size: int = None,
                         font_weight: str = 'normal',
                         font_color: str = None,
                         background_color: str = None,
                         text_align: str = None,
                         vertical_align: str = None,
                         border: str = None,
                         padding: str = None) -> Dict[str, str]:
        """创建完整的单元格样式"""
        styles = {}
        
        # 字体样式
        font_styles = self.create_font_style(font_family, font_size, font_weight, color=font_color)
        styles.update(font_styles)
        
        # 背景色
        if background_color:
            styles['background-color'] = self.get_color(background_color)
        
        # 对齐
        if text_align:
            styles['text-align'] = self.get_alignment_style(text_align)
        
        if vertical_align:
            styles['vertical-align'] = self.get_vertical_alignment_style(vertical_align)
        
        # 边框
        if border:
            styles['border'] = f'{self.get_border_style(border)} {self.DEFAULT_BORDER_COLOR}'
        
        # 内边距
        if padding:
            styles['padding'] = padding
        
        return styles
    
    def get_css_string(self, styles: Dict[str, str]) -> str:
        """将样式字典转换为CSS字符串"""
        return '; '.join([f'{key}: {value}' for key, value in styles.items()])
    
    def optimize_styles(self, styles: Dict[str, str]) -> Dict[str, str]:
        """优化样式，移除默认值"""
        optimized = {}
        
        for key, value in styles.items():
            # 跳过默认值
            if key == 'font-family' and value == self.DEFAULT_FONT_FAMILY:
                continue
            if key == 'font-size' and value == f'{self.DEFAULT_FONT_SIZE}px':
                continue
            if key == 'color' and value == self.DEFAULT_TEXT_COLOR:
                continue
            if key == 'background-color' and value == self.DEFAULT_BACKGROUND_COLOR:
                continue
            
            optimized[key] = value
        
        return optimized 