# 示例文件说明

本目录包含MCP-Sheet-Parser的各种示例文件，用于展示和测试不同格式的表格文件转换功能。

## 文件结构

### Excel格式示例
- `basic_sample.xlsx` - 基础Excel文件，包含员工信息
- `complex_sample.xlsx` - 复杂Excel文件，包含公式、样式、批注、超链接
- `template_sample.xltx` - Excel模板文件

### CSV格式示例
- `basic_sample.csv` - 基础CSV文件
- `complex_sample.csv` - 复杂CSV文件，包含特殊字符
- `multilingual_sample.csv` - 多语言CSV文件

### WPS格式示例
- `wps_sample.et` - WPS表格文件
- `wps_template.ett` - WPS模板文件

### 复杂示例
- `multi_sheet_sample.xlsx` - 多工作表Excel文件
- `merged_cells_sample.xlsx` - 包含合并单元格的Excel文件

### 图表示例
- `chart_demo.xlsx` - 包含多种图表类型的Excel文件（柱状图、饼图、折线图）

## 使用说明

1. 这些示例文件可以用于测试MCP-Sheet-Parser的各种功能
2. 每个文件都包含不同的数据结构和格式特性
3. 可以通过命令行工具转换这些文件：
   ```bash
   python main.py 示例文件/excel/basic_sample.xlsx --o output.html
   
   # 转换包含图表的文件
   python main.py 示例文件/chart_demo.xlsx --o chart_output.html --enable-charts
   ```

## 文件特性

### 基础功能
- 文本、数字、日期数据类型
- 基本的表格结构
- 简单的样式信息

### 高级功能
- 公式计算
- 单元格合并
- 批注信息
- 超链接
- 多工作表
- 复杂样式
- 图表转换

### 特殊功能
- 多语言支持
- 模板功能
- 特殊字符处理
- 大数据量处理

这些示例文件涵盖了MCP-Sheet-Parser支持的所有主要功能，是测试和演示的理想选择。
