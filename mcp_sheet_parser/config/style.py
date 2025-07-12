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
                # 标准对齐方式
                'left': 'left',
                'center': 'center',
                'right': 'right',
                'justify': 'justify',
                
                # Excel特定对齐方式
                'centerContinuous': 'center',
                'distributed': 'justify',
                'fill': 'left',
                'general': 'left',
                
                # 中文对齐方式名称支持
                '左对齐': 'left',
                '居中': 'center',
                '右对齐': 'right',
                '两端对齐': 'justify',
                '分散对齐': 'justify',
                '填充': 'left',
                '常规': 'left',
                
                # 英文别名
                'start': 'left',
                'end': 'right',
                'middle': 'center',
                'justified': 'justify',
                'distribute': 'justify',
                
                # 数字代码映射（某些Excel版本使用）
                '0': 'general',
                '1': 'left',
                '2': 'center',
                '3': 'right',
                '4': 'fill',
                '5': 'justify',
                '6': 'centerContinuous',
                '7': 'distributed'
            }
        
        if self.VERTICAL_ALIGNMENT_MAPPING is None:
            self.VERTICAL_ALIGNMENT_MAPPING = {
                # 标准垂直对齐方式
                'top': 'top',
                'center': 'middle',
                'middle': 'middle',
                'bottom': 'bottom',
                'justify': 'middle',
                'distributed': 'middle',
                
                # 中文垂直对齐方式名称支持
                '顶端对齐': 'top',
                '垂直居中': 'middle',
                '底端对齐': 'bottom',
                '垂直两端对齐': 'middle',
                '垂直分散对齐': 'middle',
                
                # 英文别名
                'start': 'top',
                'end': 'bottom',
                'baseline': 'top',
                
                # 数字代码映射
                '0': 'top',
                '1': 'center',
                '2': 'bottom',
                '3': 'justify',
                '4': 'distributed'
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
        # 处理None或空字符串
        if not alignment:
            return 'left'
        
        # 转换为字符串并标准化
        alignment_str = str(alignment).strip().lower()
        return self.ALIGNMENT_MAPPING.get(alignment_str, 'left')
    
    def get_vertical_alignment_style(self, alignment: str) -> str:
        """获取垂直对齐样式"""
        # 处理None或空字符串
        if not alignment:
            return 'top'
        
        # 转换为字符串并标准化
        alignment_str = str(alignment).strip().lower()
        return self.VERTICAL_ALIGNMENT_MAPPING.get(alignment_str, 'top')
    
    def get_color(self, color_name: str) -> str:
        """获取预设颜色"""
        return self.COLOR_PRESETS.get(color_name, color_name)
    
    def is_valid_alignment(self, alignment: str) -> bool:
        """检查对齐方式是否有效"""
        if not alignment:
            return False
        alignment_str = str(alignment).strip().lower()
        return alignment_str in self.ALIGNMENT_MAPPING
    
    def is_valid_vertical_alignment(self, alignment: str) -> bool:
        """检查垂直对齐方式是否有效"""
        if not alignment:
            return False
        alignment_str = str(alignment).strip().lower()
        return alignment_str in self.VERTICAL_ALIGNMENT_MAPPING
    
    def get_supported_alignments(self) -> list:
        """获取支持的水平对齐方式列表"""
        return list(set(self.ALIGNMENT_MAPPING.values()))
    
    def get_supported_vertical_alignments(self) -> list:
        """获取支持的垂直对齐方式列表"""
        return list(set(self.VERTICAL_ALIGNMENT_MAPPING.values()))
    
    def create_font_style(self, 
                         family: str = None, 
                         size: int = None, 
                         weight: str = 'normal',
                         style: str = 'normal',
                         color: str = None,
                         underline: bool = False,
                         strike: bool = False) -> Dict[str, str]:
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
        
        # 处理下划线和删除线
        text_decorations = []
        if underline:
            text_decorations.append('underline')
        if strike:
            text_decorations.append('line-through')
        
        if text_decorations:
            styles['text-decoration'] = ' '.join(text_decorations)
        
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
                         padding: str = None,
                         underline: bool = False,
                         strike: bool = False) -> Dict[str, str]:
        """创建完整的单元格样式"""
        styles = {}
        
        # 字体样式
        font_styles = self.create_font_style(font_family, font_size, font_weight, color=font_color, underline=underline, strike=strike)
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