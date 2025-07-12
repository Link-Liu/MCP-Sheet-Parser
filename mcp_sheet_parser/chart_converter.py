#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
图表转换模块
将Excel图表转换为SVG矢量图形
"""

import re
import math
import xml.etree.ElementTree as ET
from enum import Enum
from dataclasses import dataclass
from typing import List, Dict, Any, Optional, Tuple, Union
import logging


class ChartType(Enum):
    """图表类型枚举"""
    COLUMN = "column"           # 柱状图
    BAR = "bar"                # 条形图  
    LINE = "line"              # 折线图
    PIE = "pie"                # 饼图
    AREA = "area"              # 面积图
    SCATTER = "scatter"        # 散点图
    COMBO = "combo"            # 组合图
    UNKNOWN = "unknown"        # 未知类型


class ChartElement(Enum):
    """图表元素枚举"""
    TITLE = "title"            # 标题
    LEGEND = "legend"          # 图例
    X_AXIS = "x_axis"          # X轴
    Y_AXIS = "y_axis"          # Y轴
    DATA_SERIES = "data_series" # 数据系列
    DATA_LABELS = "data_labels" # 数据标签
    GRID_LINES = "grid_lines"   # 网格线


@dataclass
class ChartData:
    """图表数据结构"""
    chart_type: ChartType
    title: str = ""
    data_series: List[Dict[str, Any]] = None
    categories: List[str] = None
    x_axis_title: str = ""
    y_axis_title: str = ""
    legend_position: str = "right"
    colors: List[str] = None
    width: int = 600
    height: int = 400
    
    def __post_init__(self):
        if self.data_series is None:
            self.data_series = []
        if self.categories is None:
            self.categories = []
        if self.colors is None:
            self.colors = [
                "#5B9BD5", "#70AD47", "#FFC000", "#E15759", "#A5A5A5",
                "#4472C4", "#264478", "#636363", "#FF6B35", "#8E44AD"
            ]


@dataclass
class SVGElement:
    """SVG元素"""
    tag: str
    attributes: Dict[str, str]
    text: str = ""
    children: List['SVGElement'] = None
    
    def __post_init__(self):
        if self.children is None:
            self.children = []


class ChartStyler:
    """图表样式处理器"""
    
    def __init__(self):
        self.default_colors = [
            "#5B9BD5", "#70AD47", "#FFC000", "#E15759", "#A5A5A5",
            "#4472C4", "#264478", "#636363", "#FF6B35", "#8E44AD"
        ]
        
        self.chart_fonts = {
            'title': {'size': 16, 'weight': 'bold', 'family': 'Arial, sans-serif'},
            'axis': {'size': 12, 'weight': 'normal', 'family': 'Arial, sans-serif'},
            'legend': {'size': 11, 'weight': 'normal', 'family': 'Arial, sans-serif'},
            'label': {'size': 10, 'weight': 'normal', 'family': 'Arial, sans-serif'}
        }
    
    def get_color_scheme(self, chart_type: ChartType, series_count: int) -> List[str]:
        """获取图表配色方案"""
        colors = self.default_colors[:series_count]
        
        # 如果系列数超过默认颜色数，生成更多颜色
        while len(colors) < series_count:
            colors.extend(self.default_colors)
        
        return colors[:series_count]
    
    def get_font_style(self, element_type: str) -> Dict[str, str]:
        """获取字体样式"""
        return self.chart_fonts.get(element_type, self.chart_fonts['label'])


class SVGGenerator:
    """SVG生成器"""
    
    def __init__(self, width: int = 600, height: int = 400):
        self.width = width
        self.height = height
        self.margin = {'top': 50, 'right': 100, 'bottom': 80, 'left': 80}
        self.chart_area = {
            'x': self.margin['left'],
            'y': self.margin['top'],
            'width': self.width - self.margin['left'] - self.margin['right'],
            'height': self.height - self.margin['top'] - self.margin['bottom']
        }
        
    def create_svg_root(self) -> SVGElement:
        """创建SVG根元素"""
        return SVGElement(
            tag="svg",
            attributes={
                "width": str(self.width),
                "height": str(self.height),
                "viewBox": f"0 0 {self.width} {self.height}",
                "xmlns": "http://www.w3.org/2000/svg",
                "style": "background: white; font-family: Arial, sans-serif;"
            }
        )
    
    def add_title(self, svg: SVGElement, title: str) -> None:
        """添加图表标题"""
        if not title:
            return
            
        title_element = SVGElement(
            tag="text",
            attributes={
                "x": str(self.width // 2),
                "y": "30",
                "text-anchor": "middle",
                "font-size": "16",
                "font-weight": "bold",
                "fill": "#333333"
            },
            text=title
        )
        svg.children.append(title_element)
    
    def add_grid_lines(self, svg: SVGElement, x_count: int = 5, y_count: int = 5) -> None:
        """添加网格线"""
        grid_group = SVGElement(tag="g", attributes={"class": "grid-lines"})
        
        # 垂直网格线
        for i in range(x_count + 1):
            x = self.chart_area['x'] + (i * self.chart_area['width'] / x_count)
            line = SVGElement(
                tag="line",
                attributes={
                    "x1": str(x), "y1": str(self.chart_area['y']),
                    "x2": str(x), "y2": str(self.chart_area['y'] + self.chart_area['height']),
                    "stroke": "#E0E0E0", "stroke-width": "1"
                }
            )
            grid_group.children.append(line)
        
        # 水平网格线
        for i in range(y_count + 1):
            y = self.chart_area['y'] + (i * self.chart_area['height'] / y_count)
            line = SVGElement(
                tag="line",
                attributes={
                    "x1": str(self.chart_area['x']), "y1": str(y),
                    "x2": str(self.chart_area['x'] + self.chart_area['width']), "y2": str(y),
                    "stroke": "#E0E0E0", "stroke-width": "1"
                }
            )
            grid_group.children.append(line)
        
        svg.children.append(grid_group)
    
    def add_axes(self, svg: SVGElement, x_title: str = "", y_title: str = "", 
                 x_labels: List[str] = None, y_max: float = 100) -> None:
        """添加坐标轴"""
        axes_group = SVGElement(tag="g", attributes={"class": "axes"})
        
        # X轴
        x_axis = SVGElement(
            tag="line",
            attributes={
                "x1": str(self.chart_area['x']),
                "y1": str(self.chart_area['y'] + self.chart_area['height']),
                "x2": str(self.chart_area['x'] + self.chart_area['width']),
                "y2": str(self.chart_area['y'] + self.chart_area['height']),
                "stroke": "#333333", "stroke-width": "2"
            }
        )
        axes_group.children.append(x_axis)
        
        # Y轴
        y_axis = SVGElement(
            tag="line",
            attributes={
                "x1": str(self.chart_area['x']),
                "y1": str(self.chart_area['y']),
                "x2": str(self.chart_area['x']),
                "y2": str(self.chart_area['y'] + self.chart_area['height']),
                "stroke": "#333333", "stroke-width": "2"
            }
        )
        axes_group.children.append(y_axis)
        
        # X轴标签
        if x_labels:
            for i, label in enumerate(x_labels):
                x = self.chart_area['x'] + ((i + 0.5) * self.chart_area['width'] / len(x_labels))
                y = self.chart_area['y'] + self.chart_area['height'] + 20
                
                label_element = SVGElement(
                    tag="text",
                    attributes={
                        "x": str(x), "y": str(y),
                        "text-anchor": "middle",
                        "font-size": "12", "fill": "#666666"
                    },
                    text=str(label)
                )
                axes_group.children.append(label_element)
        
        # Y轴标签
        y_step = y_max / 5
        for i in range(6):
            value = i * y_step
            y = self.chart_area['y'] + self.chart_area['height'] - (i * self.chart_area['height'] / 5)
            x = self.chart_area['x'] - 10
            
            label_element = SVGElement(
                tag="text",
                attributes={
                    "x": str(x), "y": str(y + 5),
                    "text-anchor": "end",
                    "font-size": "12", "fill": "#666666"
                },
                text=f"{value:.0f}"
            )
            axes_group.children.append(label_element)
        
        # 轴标题
        if x_title:
            x_title_element = SVGElement(
                tag="text",
                attributes={
                    "x": str(self.chart_area['x'] + self.chart_area['width'] // 2),
                    "y": str(self.height - 20),
                    "text-anchor": "middle",
                    "font-size": "14", "font-weight": "bold", "fill": "#333333"
                },
                text=x_title
            )
            axes_group.children.append(x_title_element)
        
        if y_title:
            y_title_element = SVGElement(
                tag="text",
                attributes={
                    "x": "20",
                    "y": str(self.chart_area['y'] + self.chart_area['height'] // 2),
                    "text-anchor": "middle",
                    "font-size": "14", "font-weight": "bold", "fill": "#333333",
                    "transform": f"rotate(-90, 20, {self.chart_area['y'] + self.chart_area['height'] // 2})"
                },
                text=y_title
            )
            axes_group.children.append(y_title_element)
        
        svg.children.append(axes_group)
    
    def add_legend(self, svg: SVGElement, series_names: List[str], colors: List[str]) -> None:
        """添加图例"""
        if not series_names:
            return
            
        legend_group = SVGElement(tag="g", attributes={"class": "legend"})
        
        legend_x = self.chart_area['x'] + self.chart_area['width'] + 20
        legend_y = self.chart_area['y'] + 20
        
        for i, (name, color) in enumerate(zip(series_names, colors)):
            y_offset = i * 25
            
            # 颜色方块
            rect = SVGElement(
                tag="rect",
                attributes={
                    "x": str(legend_x), "y": str(legend_y + y_offset),
                    "width": "15", "height": "15",
                    "fill": color, "stroke": "#333333", "stroke-width": "1"
                }
            )
            legend_group.children.append(rect)
            
            # 标签文字
            text = SVGElement(
                tag="text",
                attributes={
                    "x": str(legend_x + 25), "y": str(legend_y + y_offset + 12),
                    "font-size": "11", "fill": "#333333"
                },
                text=name
            )
            legend_group.children.append(text)
        
        svg.children.append(legend_group)


class ColumnChartGenerator:
    """柱状图生成器"""
    
    def __init__(self, svg_generator: SVGGenerator):
        self.svg = svg_generator
    
    def generate(self, chart_data: ChartData) -> SVGElement:
        """生成柱状图SVG"""
        svg_root = self.svg.create_svg_root()
        
        # 添加标题
        self.svg.add_title(svg_root, chart_data.title)
        
        # 计算数据范围
        max_value = 0
        for series in chart_data.data_series:
            max_value = max(max_value, max(series.get('values', [0])))
        
        y_max = math.ceil(max_value * 1.1)  # 留10%空间
        
        # 添加网格线
        self.svg.add_grid_lines(svg_root)
        
        # 添加坐标轴
        self.svg.add_axes(svg_root, chart_data.x_axis_title, chart_data.y_axis_title, 
                         chart_data.categories, y_max)
        
        # 绘制柱状图
        self._draw_columns(svg_root, chart_data, y_max)
        
        # 添加图例
        series_names = [series.get('name', f'系列{i+1}') for i, series in enumerate(chart_data.data_series)]
        self.svg.add_legend(svg_root, series_names, chart_data.colors)
        
        return svg_root
    
    def _draw_columns(self, svg: SVGElement, chart_data: ChartData, y_max: float) -> None:
        """绘制柱状图的柱子"""
        columns_group = SVGElement(tag="g", attributes={"class": "columns"})
        
        category_count = len(chart_data.categories)
        series_count = len(chart_data.data_series)
        
        if category_count == 0 or series_count == 0:
            svg.children.append(columns_group)
            return
        
        category_width = self.svg.chart_area['width'] / category_count
        column_width = category_width / (series_count + 1)  # 留间距
        
        for cat_idx, category in enumerate(chart_data.categories):
            for series_idx, series in enumerate(chart_data.data_series):
                values = series.get('values', [])
                if cat_idx < len(values):
                    value = values[cat_idx]
                    
                    # 计算柱子位置和大小
                    x = (self.svg.chart_area['x'] + 
                         cat_idx * category_width + 
                         (series_idx + 0.5) * column_width)
                    
                    column_height = (value / y_max) * self.svg.chart_area['height']
                    y = self.svg.chart_area['y'] + self.svg.chart_area['height'] - column_height
                    
                    # 创建柱子
                    column = SVGElement(
                        tag="rect",
                        attributes={
                            "x": str(x), "y": str(y),
                            "width": str(column_width * 0.8),
                            "height": str(column_height),
                            "fill": chart_data.colors[series_idx % len(chart_data.colors)],
                            "stroke": "#333333", "stroke-width": "1",
                            "opacity": "0.8"
                        }
                    )
                    
                    # 添加悬停效果
                    column.attributes["style"] = "cursor: pointer;"
                    
                    # 添加标题属性显示数值
                    column.attributes["title"] = f"{series.get('name', f'系列{series_idx+1}')}: {value}"
                    
                    columns_group.children.append(column)
        
        svg.children.append(columns_group)


class PieChartGenerator:
    """饼图生成器"""
    
    def __init__(self, svg_generator: SVGGenerator):
        self.svg = svg_generator
    
    def generate(self, chart_data: ChartData) -> SVGElement:
        """生成饼图SVG"""
        svg_root = self.svg.create_svg_root()
        
        # 添加标题
        self.svg.add_title(svg_root, chart_data.title)
        
        # 绘制饼图
        self._draw_pie(svg_root, chart_data)
        
        # 添加图例
        self.svg.add_legend(svg_root, chart_data.categories, chart_data.colors)
        
        return svg_root
    
    def _draw_pie(self, svg: SVGElement, chart_data: ChartData) -> None:
        """绘制饼图"""
        pie_group = SVGElement(tag="g", attributes={"class": "pie-chart"})
        
        # 使用第一个数据系列
        if not chart_data.data_series:
            svg.children.append(pie_group)
            return
            
        values = chart_data.data_series[0].get('values', [])
        if not values:
            svg.children.append(pie_group)
            return
        
        # 计算饼图中心和半径
        center_x = self.svg.chart_area['x'] + self.svg.chart_area['width'] // 3
        center_y = self.svg.chart_area['y'] + self.svg.chart_area['height'] // 2
        radius = min(self.svg.chart_area['width'], self.svg.chart_area['height']) // 3
        
        total = sum(values)
        if total == 0:
            svg.children.append(pie_group)
            return
        
        # 绘制扇形
        start_angle = 0
        for i, (value, category) in enumerate(zip(values, chart_data.categories)):
            if value <= 0:
                continue
                
            angle = (value / total) * 360
            end_angle = start_angle + angle
            
            # 创建扇形路径
            path_data = self._create_pie_slice_path(
                center_x, center_y, radius, start_angle, end_angle
            )
            
            slice_element = SVGElement(
                tag="path",
                attributes={
                    "d": path_data,
                    "fill": chart_data.colors[i % len(chart_data.colors)],
                    "stroke": "#FFFFFF", "stroke-width": "2",
                    "title": f"{category}: {value} ({value/total*100:.1f}%)"
                }
            )
            
            pie_group.children.append(slice_element)
            
            # 添加标签
            label_angle = math.radians(start_angle + angle / 2)
            label_x = center_x + (radius * 0.7) * math.cos(label_angle)
            label_y = center_y + (radius * 0.7) * math.sin(label_angle)
            
            if value / total > 0.05:  # 只显示占比大于5%的标签
                label = SVGElement(
                    tag="text",
                    attributes={
                        "x": str(label_x), "y": str(label_y),
                        "text-anchor": "middle", "font-size": "10",
                        "fill": "white", "font-weight": "bold"
                    },
                    text=f"{value/total*100:.0f}%"
                )
                pie_group.children.append(label)
            
            start_angle = end_angle
        
        svg.children.append(pie_group)
    
    def _create_pie_slice_path(self, cx: float, cy: float, radius: float, 
                              start_angle: float, end_angle: float) -> str:
        """创建饼图扇形的SVG路径"""
        start_rad = math.radians(start_angle)
        end_rad = math.radians(end_angle)
        
        x1 = cx + radius * math.cos(start_rad)
        y1 = cy + radius * math.sin(start_rad)
        x2 = cx + radius * math.cos(end_rad)
        y2 = cy + radius * math.sin(end_rad)
        
        large_arc = 1 if (end_angle - start_angle) > 180 else 0
        
        return (f"M {cx} {cy} "
                f"L {x1} {y1} "
                f"A {radius} {radius} 0 {large_arc} 1 {x2} {y2} "
                f"Z")


class LineChartGenerator:
    """折线图生成器"""
    
    def __init__(self, svg_generator: SVGGenerator):
        self.svg = svg_generator
    
    def generate(self, chart_data: ChartData) -> SVGElement:
        """生成折线图SVG"""
        svg_root = self.svg.create_svg_root()
        
        # 添加标题
        self.svg.add_title(svg_root, chart_data.title)
        
        # 计算数据范围
        max_value = 0
        for series in chart_data.data_series:
            max_value = max(max_value, max(series.get('values', [0])))
        
        y_max = math.ceil(max_value * 1.1)
        
        # 添加网格线
        self.svg.add_grid_lines(svg_root)
        
        # 添加坐标轴
        self.svg.add_axes(svg_root, chart_data.x_axis_title, chart_data.y_axis_title,
                         chart_data.categories, y_max)
        
        # 绘制折线
        self._draw_lines(svg_root, chart_data, y_max)
        
        # 添加图例
        series_names = [series.get('name', f'系列{i+1}') for i, series in enumerate(chart_data.data_series)]
        self.svg.add_legend(svg_root, series_names, chart_data.colors)
        
        return svg_root
    
    def _draw_lines(self, svg: SVGElement, chart_data: ChartData, y_max: float) -> None:
        """绘制折线"""
        lines_group = SVGElement(tag="g", attributes={"class": "lines"})
        
        category_count = len(chart_data.categories)
        if category_count == 0:
            svg.children.append(lines_group)
            return
        
        x_step = self.svg.chart_area['width'] / (category_count - 1) if category_count > 1 else 0
        
        for series_idx, series in enumerate(chart_data.data_series):
            values = series.get('values', [])
            if not values:
                continue
            
            # 创建折线路径
            path_points = []
            for i, value in enumerate(values):
                x = self.svg.chart_area['x'] + i * x_step
                y = (self.svg.chart_area['y'] + self.svg.chart_area['height'] - 
                     (value / y_max) * self.svg.chart_area['height'])
                path_points.append(f"{x},{y}")
            
            if path_points:
                path_data = "M " + " L ".join(path_points)
                
                line = SVGElement(
                    tag="path",
                    attributes={
                        "d": path_data,
                        "fill": "none",
                        "stroke": chart_data.colors[series_idx % len(chart_data.colors)],
                        "stroke-width": "3",
                        "stroke-linecap": "round",
                        "stroke-linejoin": "round"
                    }
                )
                lines_group.children.append(line)
                
                # 添加数据点
                for i, value in enumerate(values):
                    x = self.svg.chart_area['x'] + i * x_step
                    y = (self.svg.chart_area['y'] + self.svg.chart_area['height'] - 
                         (value / y_max) * self.svg.chart_area['height'])
                    
                    point = SVGElement(
                        tag="circle",
                        attributes={
                            "cx": str(x), "cy": str(y), "r": "4",
                            "fill": chart_data.colors[series_idx % len(chart_data.colors)],
                            "stroke": "white", "stroke-width": "2",
                            "title": f"{series.get('name', f'系列{series_idx+1}')}: {value}"
                        }
                    )
                    lines_group.children.append(point)
        
        svg.children.append(lines_group)


class ChartConverter:
    """图表转换器主类"""
    
    def __init__(self, config=None, file_path=None):
        self.config = config
        self.file_path = file_path
        self.logger = logging.getLogger(__name__)
        self.styler = ChartStyler()
        
        # 图表生成器映射
        self.generators = {}
    
    def detect_charts_in_excel(self, workbook_path: str) -> List[Dict[str, Any]]:
        """
        从Excel文件中检测图表
        注意：openpyxl对图表支持有限，这里提供基础实现
        """
        charts = []
        
        try:
            import openpyxl
            wb = openpyxl.load_workbook(workbook_path)
            
            for ws in wb.worksheets:
                # openpyxl的图表检测功能有限
                # 这里先创建一些示例图表用于演示
                if hasattr(ws, '_charts') and ws._charts:
                    for chart in ws._charts:
                        chart_info = self._extract_chart_info(chart, ws)
                        if chart_info:
                            charts.append(chart_info)
        
        except Exception as e:
            self.logger.warning(f"图表检测时出现警告: {e}")
        
        return charts
    
    def _extract_chart_info(self, chart, worksheet) -> Optional[Dict[str, Any]]:
        """从Excel图表对象提取信息"""
        try:
            # 这是一个简化的实现，实际的图表提取会更复杂
            chart_info = {
                'type': ChartType.COLUMN,  # 默认类型
                'title': getattr(chart, 'title', '图表'),
                'data_series': [],
                'categories': [],
                'position': getattr(chart, 'anchor', None)
            }
            
            return chart_info
        
        except Exception as e:
            self.logger.error(f"提取图表信息失败: {e}")
            return None
    
    def create_demo_charts(self) -> List[ChartData]:
        """创建演示图表数据"""
        charts = []
        
        # 柱状图示例
        column_chart = ChartData(
            chart_type=ChartType.COLUMN,
            title="季度销售业绩对比",
            categories=["Q1", "Q2", "Q3", "Q4"],
            data_series=[
                {"name": "产品A", "values": [120, 150, 180, 200]},
                {"name": "产品B", "values": [100, 130, 160, 170]},
                {"name": "产品C", "values": [80, 110, 140, 150]}
            ],
            x_axis_title="季度",
            y_axis_title="销售额(万元)",
            width=600,
            height=400
        )
        charts.append(column_chart)
        
        # 饼图示例
        pie_chart = ChartData(
            chart_type=ChartType.PIE,
            title="市场份额分布",
            categories=["华东", "华南", "华北", "华中", "其他"],
            data_series=[
                {"name": "市场份额", "values": [35, 25, 20, 15, 5]}
            ],
            width=500,
            height=400
        )
        charts.append(pie_chart)
        
        # 折线图示例
        line_chart = ChartData(
            chart_type=ChartType.LINE,
            title="月度增长趋势",
            categories=["1月", "2月", "3月", "4月", "5月", "6月"],
            data_series=[
                {"name": "收入", "values": [100, 120, 110, 140, 160, 180]},
                {"name": "利润", "values": [20, 25, 22, 30, 35, 40]}
            ],
            x_axis_title="月份",
            y_axis_title="金额(万元)",
            width=600,
            height=400
        )
        charts.append(line_chart)
        
        return charts
    
    def generate_svg(self, chart_data: ChartData) -> str:
        """生成SVG图表"""
        try:
            # 创建SVG生成器
            svg_generator = SVGGenerator(chart_data.width, chart_data.height)
            
            # 根据图表类型选择生成器
            if chart_data.chart_type == ChartType.COLUMN:
                generator = ColumnChartGenerator(svg_generator)
            elif chart_data.chart_type == ChartType.PIE:
                generator = PieChartGenerator(svg_generator)
            elif chart_data.chart_type == ChartType.LINE:
                generator = LineChartGenerator(svg_generator)
            else:
                # 默认使用柱状图
                generator = ColumnChartGenerator(svg_generator)
            
            # 生成SVG元素
            svg_element = generator.generate(chart_data)
            
            # 转换为字符串
            return self._svg_element_to_string(svg_element)
            
        except Exception as e:
            self.logger.error(f"生成SVG图表失败: {e}")
            return self._create_error_svg(str(e))
    
    def _svg_element_to_string(self, element: SVGElement) -> str:
        """将SVG元素转换为字符串"""
        # 构建属性字符串
        attrs = ' '.join([f'{k}="{v}"' for k, v in element.attributes.items()])
        
        # 处理子元素
        children_str = ""
        for child in element.children:
            children_str += self._svg_element_to_string(child)
        
        # 构建完整标签
        if element.text or children_str:
            return f'<{element.tag} {attrs}>{element.text}{children_str}</{element.tag}>'
        else:
            return f'<{element.tag} {attrs}/>'
    
    def _create_error_svg(self, error_message: str) -> str:
        """创建错误提示SVG"""
        return f"""
        <svg width="400" height="200" xmlns="http://www.w3.org/2000/svg">
            <rect width="400" height="200" fill="#f8f9fa" stroke="#dee2e6"/>
            <text x="200" y="100" text-anchor="middle" font-family="Arial" font-size="14" fill="#dc3545">
                图表生成失败: {error_message}
            </text>
        </svg>
        """
    
    def process_sheet_charts(self, sheet_data: Dict[str, Any]) -> Dict[str, Any]:
        """处理工作表中的图表"""
        try:
            # 首先尝试检测真实的图表
            real_charts = self.detect_charts_in_excel(self.file_path)
            
            # 如果没有检测到真实图表，则不创建演示图表
            if not real_charts:
                self.logger.info("未检测到真实图表，跳过图表处理")
                return sheet_data
            
            # 如果检测到真实图表，则处理它们
            charts = real_charts
            
            # 生成SVG
            chart_svgs = []
            for chart_data in charts:
                svg_content = self.generate_svg(chart_data)
                chart_svgs.append({
                    'type': chart_data.chart_type.value,
                    'title': chart_data.title,
                    'svg': svg_content,
                    'width': chart_data.width,
                    'height': chart_data.height
                })
            
            # 添加图表信息到工作表数据
            enhanced_data = sheet_data.copy()
            enhanced_data['charts'] = chart_svgs
            
            self.logger.info(f"处理了 {len(chart_svgs)} 个真实图表")
            return enhanced_data
            
        except Exception as e:
            self.logger.error(f"处理图表时出错: {e}")
            return sheet_data 