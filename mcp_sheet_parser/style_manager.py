# style_manager.py
# 高级样式管理器 - CSS类生成与条件格式化

import hashlib
import re
import logging
from typing import Dict, List, Tuple, Any, Optional, Union
from dataclasses import dataclass, field
from enum import Enum


class ConditionalType(Enum):
    """条件格式化类型"""
    VALUE_RANGE = "value_range"
    COLOR_SCALE = "color_scale" 
    DATA_BARS = "data_bars"
    ICON_SETS = "icon_sets"
    DUPLICATE_VALUES = "duplicate_values"
    TOP_BOTTOM = "top_bottom"
    CUSTOM_FORMULA = "custom_formula"


class ComparisonOperator(Enum):
    """比较操作符"""
    GREATER_THAN = ">"
    LESS_THAN = "<"
    GREATER_EQUAL = ">="
    LESS_EQUAL = "<="
    EQUAL = "=="
    NOT_EQUAL = "!="
    BETWEEN = "between"
    NOT_BETWEEN = "not_between"
    CONTAINS = "contains"
    NOT_CONTAINS = "not_contains"


@dataclass
class ConditionalRule:
    """条件格式化规则"""
    name: str
    type: ConditionalType
    operator: ComparisonOperator
    values: List[Any]
    styles: Dict[str, Any]
    priority: int = 0
    enabled: bool = True
    description: str = ""


@dataclass
class StyleClass:
    """CSS类定义"""
    name: str
    css_properties: Dict[str, str]
    usage_count: int = 0
    hash_key: str = ""
    is_semantic: bool = False
    description: str = ""


@dataclass
class StyleTemplate:
    """样式模板"""
    name: str
    description: str
    classes: Dict[str, StyleClass] = field(default_factory=dict)
    conditional_rules: List[ConditionalRule] = field(default_factory=list)
    global_css: str = ""


class StyleManager:
    """高级样式管理器"""
    
    def __init__(self, config=None):
        from .config import Config
        self.config = config or Config()
        self.logger = logging.getLogger(__name__)
        
        # 样式类存储
        self.style_classes: Dict[str, StyleClass] = {}
        self.style_hash_map: Dict[str, str] = {}  # hash -> class_name
        
        # 条件格式化规则
        self.conditional_rules: List[ConditionalRule] = []
        
        # 当前使用的模板
        self.current_template: Optional[StyleTemplate] = None
        
        # 预定义模板
        self.templates = self._init_default_templates()
        
        # 统计信息
        self.stats = {
            'total_styles': 0,
            'unique_styles': 0,
            'class_reuse_rate': 0.0,
            'conditional_rules_applied': 0
        }

    def _init_default_templates(self) -> Dict[str, StyleTemplate]:
        """初始化默认样式模板"""
        templates = {}
        
        # 商务模板
        business_template = StyleTemplate(
            name="business",
            description="商务报表样式模板",
            global_css="""
            .business-table { 
                font-family: 'Segoe UI', Arial, sans-serif;
                border-collapse: collapse;
                width: 100%;
                margin: 20px 0;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }
            .business-header {
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                font-weight: 600;
                text-align: center;
                padding: 12px;
                border: none;
            }
            .business-row:nth-child(even) {
                background-color: #f8f9fa;
            }
            .business-row:hover {
                background-color: #e9ecef;
                transition: background-color 0.3s ease;
            }
            """
        )
        templates["business"] = business_template
        
        # 财务模板
        financial_template = StyleTemplate(
            name="financial",
            description="财务报表样式模板",
            global_css="""
            .financial-table {
                font-family: 'Times New Roman', serif;
                border-collapse: collapse;
                width: 100%;
                background: white;
            }
            .financial-number {
                text-align: right;
                font-family: 'Courier New', monospace;
                font-weight: 500;
            }
            .financial-positive {
                color: #28a745;
            }
            .financial-negative {
                color: #dc3545;
            }
            .financial-header {
                background-color: #2c3e50;
                color: white;
                font-weight: bold;
                text-align: center;
                padding: 10px;
                border: 1px solid #34495e;
            }
            """
        )
        templates["financial"] = financial_template
        
        # 数据分析模板
        analytics_template = StyleTemplate(
            name="analytics",
            description="数据分析样式模板",
            global_css="""
            .analytics-table {
                font-family: 'Roboto', sans-serif;
                border-collapse: collapse;
                width: 100%;
                font-size: 14px;
            }
            .analytics-metric {
                font-weight: 600;
                text-align: center;
                padding: 8px;
            }
            .analytics-highlight {
                background: linear-gradient(90deg, #ff6b6b, #ffa500);
                color: white;
                font-weight: bold;
            }
            .analytics-low {
                background-color: #ffebee;
                color: #c62828;
            }
            .analytics-medium {
                background-color: #fff3e0;
                color: #ef6c00;
            }
            .analytics-high {
                background-color: #e8f5e8;
                color: #2e7d32;
            }
            """
        )
        templates["analytics"] = analytics_template
        
        return templates

    def generate_css_classes(self, sheet_data: Dict[str, Any], 
                           use_semantic_names: bool = True,
                           min_usage_threshold: int = 2) -> Tuple[str, Dict[str, str]]:
        """
        生成CSS类
        
        Args:
            sheet_data: 工作表数据
            use_semantic_names: 使用语义化类名
            min_usage_threshold: 最小使用次数阈值
            
        Returns:
            (CSS样式表, 单元格类名映射)
        """
        self.logger.info("开始生成CSS类...")
        
        # 重置统计
        self.style_classes.clear()
        self.style_hash_map.clear()
        
        styles = sheet_data.get('styles', [])
        data = sheet_data.get('data', [])
        
        if not styles:
            return "", {}
        
        # 第一遍：收集所有样式，计算使用频率
        style_usage = {}
        cell_style_map = {}
        
        for row_idx, row_styles in enumerate(styles):
            for col_idx, cell_style in enumerate(row_styles):
                if cell_style:
                    # 计算样式哈希
                    style_hash = self._calculate_style_hash(cell_style)
                    
                    # 记录使用频率
                    if style_hash not in style_usage:
                        style_usage[style_hash] = {
                            'count': 0,
                            'style': cell_style,
                            'positions': []
                        }
                    
                    style_usage[style_hash]['count'] += 1
                    style_usage[style_hash]['positions'].append((row_idx, col_idx))
                    cell_style_map[(row_idx, col_idx)] = style_hash
        
        # 第二遍：生成CSS类
        cell_class_map = {}
        
        for style_hash, usage_info in style_usage.items():
            if usage_info['count'] >= min_usage_threshold:
                # 生成类名
                if use_semantic_names:
                    class_name = self._generate_semantic_class_name(
                        usage_info['style'], data, usage_info['positions']
                    )
                else:
                    class_name = f"style-{style_hash[:8]}"
                
                # 创建样式类
                style_class = StyleClass(
                    name=class_name,
                    css_properties=self._convert_style_to_css_dict(usage_info['style']),
                    usage_count=usage_info['count'],
                    hash_key=style_hash,
                    is_semantic=use_semantic_names,
                    description=self._generate_style_description(usage_info['style'])
                )
                
                self.style_classes[class_name] = style_class
                self.style_hash_map[style_hash] = class_name
                
                # 映射单元格到类名
                for pos in usage_info['positions']:
                    cell_class_map[pos] = class_name
        
        # 生成CSS样式表
        css_content = self._generate_css_stylesheet()
        
        # 更新统计信息
        self._update_statistics(style_usage, min_usage_threshold)
        
        self.logger.info(f"CSS类生成完成：{len(self.style_classes)}个类，复用率{self.stats['class_reuse_rate']:.1f}%")
        
        return css_content, cell_class_map

    def apply_conditional_formatting(self, sheet_data: Dict[str, Any], 
                                   rules: Optional[List[ConditionalRule]] = None) -> Dict[str, str]:
        """
        应用条件格式化
        
        Args:
            sheet_data: 工作表数据
            rules: 条件格式化规则（可选）
            
        Returns:
            单元格条件样式映射
        """
        if rules is None:
            rules = self.conditional_rules
        
        if not rules:
            return {}
        
        self.logger.info(f"开始应用条件格式化：{len(rules)}个规则")
        
        data = sheet_data.get('data', [])
        conditional_styles = {}
        applied_count = 0
        
        for rule in rules:
            if not rule.enabled:
                continue
            
            rule_styles = self._apply_conditional_rule(rule, data)
            
            # 合并样式（按优先级）
            for cell_pos, style in rule_styles.items():
                if cell_pos not in conditional_styles:
                    conditional_styles[cell_pos] = style
                    applied_count += 1
                else:
                    # 比较优先级
                    existing_priority = getattr(conditional_styles[cell_pos], 'priority', 0)
                    if rule.priority > existing_priority:
                        conditional_styles[cell_pos] = style
                        applied_count += 1
        
        self.stats['conditional_rules_applied'] = applied_count
        self.logger.info(f"条件格式化完成：{applied_count}个单元格应用了条件样式")
        
        return conditional_styles

    def _apply_conditional_rule(self, rule: ConditionalRule, data: List[List[Any]]) -> Dict[Tuple[int, int], str]:
        """应用单个条件格式化规则"""
        result = {}
        
        if rule.type == ConditionalType.VALUE_RANGE:
            result = self._apply_value_range_rule(rule, data)
        elif rule.type == ConditionalType.COLOR_SCALE:
            result = self._apply_color_scale_rule(rule, data)
        elif rule.type == ConditionalType.DATA_BARS:
            result = self._apply_data_bars_rule(rule, data)
        elif rule.type == ConditionalType.TOP_BOTTOM:
            result = self._apply_top_bottom_rule(rule, data)
        elif rule.type == ConditionalType.DUPLICATE_VALUES:
            result = self._apply_duplicate_values_rule(rule, data)
        
        return result

    def _apply_value_range_rule(self, rule: ConditionalRule, data: List[List[Any]]) -> Dict[Tuple[int, int], str]:
        """应用数值范围条件"""
        result = {}
        
        for row_idx, row in enumerate(data):
            for col_idx, cell_value in enumerate(row):
                if self._is_numeric(cell_value):
                    value = float(cell_value)
                    
                    if self._check_value_condition(value, rule.operator, rule.values):
                        css_style = self._dict_to_css_string(rule.styles)
                        result[(row_idx, col_idx)] = css_style
        
        return result

    def _apply_color_scale_rule(self, rule: ConditionalRule, data: List[List[Any]]) -> Dict[Tuple[int, int], str]:
        """应用颜色渐变条件"""
        result = {}
        
        # 收集所有数值
        numeric_values = []
        value_positions = []
        
        for row_idx, row in enumerate(data):
            for col_idx, cell_value in enumerate(row):
                if self._is_numeric(cell_value):
                    numeric_values.append(float(cell_value))
                    value_positions.append((row_idx, col_idx))
        
        if not numeric_values:
            return result
        
        min_val = min(numeric_values)
        max_val = max(numeric_values)
        
        # 生成颜色渐变
        colors = rule.values if len(rule.values) >= 2 else ['#ff0000', '#00ff00']
        
        for i, (row_idx, col_idx) in enumerate(value_positions):
            value = numeric_values[i]
            
            # 计算颜色插值
            if max_val != min_val:
                ratio = (value - min_val) / (max_val - min_val)
            else:
                ratio = 0.5
            
            color = self._interpolate_color(colors[0], colors[1], ratio)
            
            css_style = f"background-color: {color};"
            result[(row_idx, col_idx)] = css_style
        
        return result

    def _apply_data_bars_rule(self, rule: ConditionalRule, data: List[List[Any]]) -> Dict[Tuple[int, int], str]:
        """应用数据条条件"""
        result = {}
        
        # 收集数值数据
        numeric_values = []
        value_positions = []
        
        for row_idx, row in enumerate(data):
            for col_idx, cell_value in enumerate(row):
                if self._is_numeric(cell_value):
                    numeric_values.append(float(cell_value))
                    value_positions.append((row_idx, col_idx))
        
        if not numeric_values:
            return result
        
        min_val = min(numeric_values)
        max_val = max(numeric_values)
        
        bar_color = rule.values[0] if rule.values else '#4472C4'
        
        for i, (row_idx, col_idx) in enumerate(value_positions):
            value = numeric_values[i]
            
            # 计算条形宽度百分比
            if max_val != min_val:
                width_percent = ((value - min_val) / (max_val - min_val)) * 100
            else:
                width_percent = 50
            
            css_style = f"""
                background: linear-gradient(to right, 
                    {bar_color} 0%, 
                    {bar_color} {width_percent}%, 
                    transparent {width_percent}%, 
                    transparent 100%);
                padding: 4px;
            """.replace('\n', ' ').strip()
            
            result[(row_idx, col_idx)] = css_style
        
        return result

    def _apply_top_bottom_rule(self, rule: ConditionalRule, data: List[List[Any]]) -> Dict[Tuple[int, int], str]:
        """应用前N个/后N个值条件"""
        result = {}
        
        # 收集数值数据
        numeric_data = []
        
        for row_idx, row in enumerate(data):
            for col_idx, cell_value in enumerate(row):
                if self._is_numeric(cell_value):
                    numeric_data.append((float(cell_value), row_idx, col_idx))
        
        if not numeric_data:
            return result
        
        # 排序
        numeric_data.sort(key=lambda x: x[0], reverse=True)
        
        count = int(rule.values[0]) if rule.values else 10
        is_top = rule.operator == ComparisonOperator.GREATER_THAN
        
        selected_indices = range(min(count, len(numeric_data)))
        if not is_top:
            selected_indices = range(max(0, len(numeric_data) - count), len(numeric_data))
        
        css_style = self._dict_to_css_string(rule.styles)
        
        for i in selected_indices:
            _, row_idx, col_idx = numeric_data[i]
            result[(row_idx, col_idx)] = css_style
        
        return result

    def _apply_duplicate_values_rule(self, rule: ConditionalRule, data: List[List[Any]]) -> Dict[Tuple[int, int], str]:
        """应用重复值条件"""
        result = {}
        
        # 统计值出现次数
        value_counts = {}
        value_positions = {}
        
        for row_idx, row in enumerate(data):
            for col_idx, cell_value in enumerate(row):
                if cell_value is not None and str(cell_value).strip():
                    key = str(cell_value).strip().lower()
                    
                    if key not in value_counts:
                        value_counts[key] = 0
                        value_positions[key] = []
                    
                    value_counts[key] += 1
                    value_positions[key].append((row_idx, col_idx))
        
        css_style = self._dict_to_css_string(rule.styles)
        
        # 标记重复值
        for key, count in value_counts.items():
            if count > 1:  # 重复值
                for row_idx, col_idx in value_positions[key]:
                    result[(row_idx, col_idx)] = css_style
        
        return result

    def add_conditional_rule(self, rule: ConditionalRule):
        """添加条件格式化规则"""
        self.conditional_rules.append(rule)
        self.logger.info(f"添加条件格式化规则：{rule.name}")

    def remove_conditional_rule(self, rule_name: str) -> bool:
        """移除条件格式化规则"""
        for i, rule in enumerate(self.conditional_rules):
            if rule.name == rule_name:
                del self.conditional_rules[i]
                self.logger.info(f"移除条件格式化规则：{rule_name}")
                return True
        return False

    def set_template(self, template_name: str) -> bool:
        """设置样式模板"""
        if template_name in self.templates:
            self.current_template = self.templates[template_name]
            self.conditional_rules = self.current_template.conditional_rules.copy()
            self.logger.info(f"设置样式模板：{template_name}")
            return True
        return False

    def create_custom_template(self, template: StyleTemplate):
        """创建自定义模板"""
        self.templates[template.name] = template
        self.logger.info(f"创建自定义模板：{template.name}")

    def get_style_statistics(self) -> Dict[str, Any]:
        """获取样式统计信息"""
        return self.stats.copy()

    def _calculate_style_hash(self, style: Dict[str, Any]) -> str:
        """计算样式哈希值"""
        # 创建一个稳定的字符串表示
        style_str = str(sorted(style.items()))
        return hashlib.md5(style_str.encode()).hexdigest()

    def _generate_semantic_class_name(self, style: Dict[str, Any], 
                                    data: List[List[Any]], 
                                    positions: List[Tuple[int, int]]) -> str:
        """生成语义化类名"""
        # 分析样式特征
        features = []
        
        # 字体样式
        if style.get('bold'):
            features.append('bold')
        if style.get('italic'):
            features.append('italic')
        
        # 颜色特征
        if style.get('font_color'):
            color = style['font_color'].lower()
            if 'red' in color or color.startswith('#ff') or color.startswith('#dc'):
                features.append('red')
            elif 'green' in color or color.startswith('#00ff') or color.startswith('#28'):
                features.append('green')
            elif 'blue' in color or color.startswith('#0000ff') or color.startswith('#007'):
                features.append('blue')
        
        # 背景色特征
        if style.get('bg_color'):
            bg = style['bg_color'].lower()
            if 'yellow' in bg or bg.startswith('#ffff'):
                features.append('highlight')
            elif bg.startswith('#f'):
                features.append('light')
        
        # 位置特征
        if positions:
            row_indices = [pos[0] for pos in positions]
            if all(r == 0 for r in row_indices):
                features.append('header')
            elif len(set(row_indices)) == 1:
                features.append('row')
            elif len(set(pos[1] for pos in positions)) == 1:
                features.append('column')
        
        # 数据类型特征
        if positions and data:
            sample_values = []
            for row_idx, col_idx in positions[:5]:  # 采样前5个
                if row_idx < len(data) and col_idx < len(data[row_idx]):
                    sample_values.append(data[row_idx][col_idx])
            
            if sample_values:
                if all(self._is_numeric(v) for v in sample_values if v is not None):
                    features.append('number')
                elif any(self._is_date_like(v) for v in sample_values if v is not None):
                    features.append('date')
        
        # 生成类名
        if features:
            class_name = '-'.join(features[:3])  # 最多3个特征
        else:
            class_name = f'style-{hash(str(style)) % 10000:04d}'
        
        # 确保类名唯一
        base_name = class_name
        counter = 1
        while class_name in self.style_classes:
            class_name = f"{base_name}-{counter}"
            counter += 1
        
        return class_name

    def _convert_style_to_css_dict(self, style: Dict[str, Any]) -> Dict[str, str]:
        """将样式转换为CSS属性字典"""
        css_props = {}
        
        # 字体样式
        if style.get('bold'):
            css_props['font-weight'] = 'bold'
        if style.get('italic'):
            css_props['font-style'] = 'italic'
        if style.get('underline'):
            css_props['text-decoration'] = 'underline'
        if style.get('strike'):
            css_props['text-decoration'] = 'line-through'
        
        # 字体属性
        if style.get('font_size'):
            css_props['font-size'] = f"{style['font_size']}pt"
        if style.get('font_name'):
            css_props['font-family'] = f'"{style["font_name"]}"'
        if style.get('font_color'):
            css_props['color'] = style['font_color']
        
        # 背景和对齐
        if style.get('bg_color'):
            css_props['background-color'] = style['bg_color']
        if style.get('align'):
            css_props['text-align'] = style['align']
        if style.get('valign'):
            valign_map = {'top': 'top', 'center': 'middle', 'bottom': 'bottom'}
            if style['valign'] in valign_map:
                css_props['vertical-align'] = valign_map[style['valign']]
        
        # 边框
        if style.get('border'):
            border_css = self._convert_border_to_css_dict(style['border'])
            css_props.update(border_css)
        
        # 文本换行
        if style.get('wrap_text'):
            css_props['white-space'] = 'pre-wrap'
        
        return css_props

    def _convert_border_to_css_dict(self, border: Dict[str, Any]) -> Dict[str, str]:
        """将边框转换为CSS属性字典"""
        css_props = {}
        
        for side in ['top', 'bottom', 'left', 'right']:
            if side in border:
                border_info = border[side]
                style_name = border_info.get('style', 'solid')
                color = border_info.get('color', '#000000')
                
                # 转换样式
                style_map = {
                    'thin': '1px solid',
                    'thick': '3px solid',
                    'medium': '2px solid',
                    'dashed': '1px dashed',
                    'dotted': '1px dotted',
                    'double': '3px double'
                }
                
                css_style = style_map.get(style_name, f'{style_name} 1px')
                css_props[f'border-{side}'] = f'{css_style} {color}'
        
        return css_props

    def _generate_css_stylesheet(self) -> str:
        """生成CSS样式表"""
        css_parts = []
        
        # 添加模板全局CSS
        if self.current_template and self.current_template.global_css:
            css_parts.append("/* 模板全局样式 */")
            css_parts.append(self.current_template.global_css)
            css_parts.append("")
        
        # 添加生成的类样式
        if self.style_classes:
            css_parts.append("/* 生成的CSS类 */")
            
            for class_name, style_class in self.style_classes.items():
                # 添加注释
                if style_class.description:
                    css_parts.append(f"/* {style_class.description} (使用{style_class.usage_count}次) */")
                
                # 生成CSS规则
                properties = []
                for prop, value in style_class.css_properties.items():
                    properties.append(f"    {prop}: {value};")
                
                css_parts.append(f".{class_name} {{")
                css_parts.extend(properties)
                css_parts.append("}")
                css_parts.append("")
        
        return '\n'.join(css_parts)

    def _generate_style_description(self, style: Dict[str, Any]) -> str:
        """生成样式描述"""
        features = []
        
        if style.get('bold'):
            features.append("粗体")
        if style.get('italic'):
            features.append("斜体")
        if style.get('font_color'):
            features.append(f"文字颜色{style['font_color']}")
        if style.get('bg_color'):
            features.append(f"背景色{style['bg_color']}")
        if style.get('font_size'):
            features.append(f"{style['font_size']}pt")
        
        return ", ".join(features) if features else "自定义样式"

    def _update_statistics(self, style_usage: Dict[str, Any], min_threshold: int):
        """更新统计信息"""
        total_usages = sum(info['count'] for info in style_usage.values())
        unique_styles = len(style_usage)
        reusable_styles = len([info for info in style_usage.values() if info['count'] >= min_threshold])
        reused_usages = sum(info['count'] for info in style_usage.values() if info['count'] >= min_threshold)
        
        self.stats.update({
            'total_styles': total_usages,
            'unique_styles': unique_styles,
            'reusable_styles': reusable_styles,
            'class_reuse_rate': (reused_usages / total_usages * 100) if total_usages > 0 else 0
        })

    def _check_value_condition(self, value: float, operator: ComparisonOperator, values: List[Any]) -> bool:
        """检查数值条件"""
        if not values:
            return False
        
        if operator == ComparisonOperator.GREATER_THAN:
            return value > float(values[0])
        elif operator == ComparisonOperator.LESS_THAN:
            return value < float(values[0])
        elif operator == ComparisonOperator.GREATER_EQUAL:
            return value >= float(values[0])
        elif operator == ComparisonOperator.LESS_EQUAL:
            return value <= float(values[0])
        elif operator == ComparisonOperator.EQUAL:
            return abs(value - float(values[0])) < 1e-9
        elif operator == ComparisonOperator.BETWEEN:
            return len(values) >= 2 and float(values[0]) <= value <= float(values[1])
        elif operator == ComparisonOperator.NOT_BETWEEN:
            return len(values) >= 2 and not (float(values[0]) <= value <= float(values[1]))
        
        return False

    def _is_numeric(self, value: Any) -> bool:
        """检查是否为数值"""
        if value is None:
            return False
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False

    def _is_date_like(self, value: Any) -> bool:
        """检查是否像日期"""
        if not isinstance(value, str):
            return False
        
        date_patterns = [
            r'\d{4}-\d{2}-\d{2}',
            r'\d{2}/\d{2}/\d{4}',
            r'\d{4}年\d{1,2}月\d{1,2}日'
        ]
        
        for pattern in date_patterns:
            if re.search(pattern, value):
                return True
        
        return False

    def _interpolate_color(self, color1: str, color2: str, ratio: float) -> str:
        """颜色插值"""
        def hex_to_rgb(hex_color):
            hex_color = hex_color.lstrip('#')
            return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        def rgb_to_hex(rgb):
            return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        
        try:
            rgb1 = hex_to_rgb(color1)
            rgb2 = hex_to_rgb(color2)
            
            interpolated = tuple(
                int(rgb1[i] + (rgb2[i] - rgb1[i]) * ratio)
                for i in range(3)
            )
            
            return rgb_to_hex(interpolated)
        except:
            return color1

    def _dict_to_css_string(self, style_dict: Dict[str, Any]) -> str:
        """将样式字典转换为CSS字符串"""
        css_parts = []
        for key, value in style_dict.items():
            css_parts.append(f"{key}: {value}")
        return "; ".join(css_parts)


# 预定义条件格式化规则
PREDEFINED_RULES = {
    'financial_positive': ConditionalRule(
        name="财务正值",
        type=ConditionalType.VALUE_RANGE,
        operator=ComparisonOperator.GREATER_THAN,
        values=[0],
        styles={'color': '#28a745', 'font-weight': 'bold'},
        description="突出显示正数值"
    ),
    
    'financial_negative': ConditionalRule(
        name="财务负值",
        type=ConditionalType.VALUE_RANGE,
        operator=ComparisonOperator.LESS_THAN,
        values=[0],
        styles={'color': '#dc3545', 'font-weight': 'bold'},
        description="突出显示负数值"
    ),
    
    'top_10_percent': ConditionalRule(
        name="前10%",
        type=ConditionalType.TOP_BOTTOM,
        operator=ComparisonOperator.GREATER_THAN,
        values=[0.1],  # 10%
        styles={'background-color': '#d4edda', 'color': '#155724'},
        description="突出显示前10%的值"
    ),
    
    'data_bars_blue': ConditionalRule(
        name="蓝色数据条",
        type=ConditionalType.DATA_BARS,
        operator=ComparisonOperator.GREATER_THAN,
        values=['#4472C4'],
        styles={},
        description="蓝色数据条显示"
    ),
    
    'color_scale_red_green': ConditionalRule(
        name="红绿色阶",
        type=ConditionalType.COLOR_SCALE,
        operator=ComparisonOperator.BETWEEN,
        values=['#ff0000', '#00ff00'],
        styles={},
        description="红色到绿色的颜色渐变"
    )
} 