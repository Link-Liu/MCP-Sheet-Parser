# html.py
# HTML输出相关配置

from dataclasses import dataclass
from typing import Dict


@dataclass
class HTMLConfig:
    """HTML输出配置"""
    
    # HTML输出设置
    HTML_DEFAULT_ENCODING: str = 'utf-8'
    HTML_TABLE_BORDER: int = 1
    HTML_CELL_SPACING: int = 0
    HTML_CELL_PADDING: int = 4
    
    # 输出选项
    INCLUDE_EMPTY_CELLS: bool = True
    INCLUDE_COMMENTS: bool = True
    INCLUDE_HYPERLINKS: bool = True
    INCLUDE_MERGED_CELLS: bool = True
    INCLUDE_STYLES: bool = True
    
    # 压缩选项
    MINIFY_HTML: bool = False
    COMPRESS_INLINE_STYLES: bool = True
    REMOVE_EMPTY_ATTRIBUTES: bool = True
    
    # 兼容性选项
    HTML5_COMPATIBLE: bool = True
    IE_COMPATIBLE: bool = False
    MOBILE_FRIENDLY: bool = True
    
    # 响应式设计
    RESPONSIVE_TABLES: bool = True
    BREAKPOINT_MOBILE: str = '768px'
    BREAKPOINT_TABLET: str = '1024px'
    
    def get_html_attributes(self) -> Dict[str, str]:
        """获取HTML表格属性"""
        attributes = {
            'border': str(self.HTML_TABLE_BORDER),
            'cellspacing': str(self.HTML_CELL_SPACING),
            'cellpadding': str(self.HTML_CELL_PADDING)
        }
        
        if not self.HTML5_COMPATIBLE:
            # 为老版本浏览器添加额外属性
            attributes.update({
                'role': 'table',
                'aria-label': '表格数据'
            })
        
        return attributes
    
    def get_content_options(self) -> Dict[str, bool]:
        """获取内容包含选项"""
        return {
            'include_empty_cells': self.INCLUDE_EMPTY_CELLS,
            'include_comments': self.INCLUDE_COMMENTS,
            'include_hyperlinks': self.INCLUDE_HYPERLINKS,
            'include_merged_cells': self.INCLUDE_MERGED_CELLS,
            'include_styles': self.INCLUDE_STYLES
        }
    
    def get_meta_tags(self) -> list[str]:
        """获取HTML元标签"""
        meta_tags = [
            f'<meta charset="{self.HTML_DEFAULT_ENCODING}">',
            '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
        ]
        
        if self.MOBILE_FRIENDLY:
            meta_tags.append('<meta name="mobile-web-app-capable" content="yes">')
        
        if self.IE_COMPATIBLE:
            meta_tags.append('<meta http-equiv="X-UA-Compatible" content="IE=edge">')
        
        return meta_tags
    
    def get_responsive_css(self) -> str:
        """获取响应式CSS"""
        if not self.RESPONSIVE_TABLES:
            return ""
        
        return f"""
        @media (max-width: {self.BREAKPOINT_MOBILE}) {{
            table {{ 
                font-size: 12px; 
                display: block;
                overflow-x: auto;
                white-space: nowrap;
            }}
            td, th {{ 
                padding: 4px 2px; 
                min-width: 80px;
            }}
        }}
        
        @media (max-width: {self.BREAKPOINT_TABLET}) {{
            table {{ 
                font-size: 14px; 
            }}
            td, th {{ 
                padding: 6px 4px; 
            }}
        }}
        """
    
    def should_include_content(self, content_type: str) -> bool:
        """检查是否应该包含特定类型的内容"""
        content_map = {
            'empty_cells': self.INCLUDE_EMPTY_CELLS,
            'comments': self.INCLUDE_COMMENTS,
            'hyperlinks': self.INCLUDE_HYPERLINKS,
            'merged_cells': self.INCLUDE_MERGED_CELLS,
            'styles': self.INCLUDE_STYLES
        }
        return content_map.get(content_type, True)
    
    def optimize_for_size(self) -> 'HTMLConfig':
        """优化配置以减小输出文件大小"""
        optimized = HTMLConfig(
            HTML_DEFAULT_ENCODING=self.HTML_DEFAULT_ENCODING,
            HTML_TABLE_BORDER=self.HTML_TABLE_BORDER,
            HTML_CELL_SPACING=self.HTML_CELL_SPACING,
            HTML_CELL_PADDING=self.HTML_CELL_PADDING,
            INCLUDE_EMPTY_CELLS=False,  # 不包含空单元格
            INCLUDE_COMMENTS=False,     # 不包含注释
            INCLUDE_HYPERLINKS=self.INCLUDE_HYPERLINKS,
            INCLUDE_MERGED_CELLS=self.INCLUDE_MERGED_CELLS,
            INCLUDE_STYLES=True,        # 保留样式
            MINIFY_HTML=True,          # 启用压缩
            COMPRESS_INLINE_STYLES=True,
            REMOVE_EMPTY_ATTRIBUTES=True,
            HTML5_COMPATIBLE=self.HTML5_COMPATIBLE,
            IE_COMPATIBLE=self.IE_COMPATIBLE,
            MOBILE_FRIENDLY=self.MOBILE_FRIENDLY,
            RESPONSIVE_TABLES=self.RESPONSIVE_TABLES,
            BREAKPOINT_MOBILE=self.BREAKPOINT_MOBILE,
            BREAKPOINT_TABLET=self.BREAKPOINT_TABLET
        )
        
        return optimized 