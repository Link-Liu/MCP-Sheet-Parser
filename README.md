# MCP-Sheet-Parser

一个专注于HTML转换的表格解析工具，支持Excel、CSV、WPS等格式转换为美观的HTML表格。

## ✨ 特性

- 🚀 **快速转换**: 高效解析Excel、CSV、WPS等表格文件
- 🎨 **多种主题**: 内置4种精美主题（默认、极简、暗色、打印）
- 📊 **完整功能**: 支持合并单元格、样式、批注、超链接、公式结构化输出
- 📈 **图表转换**: 支持Excel图表转换为SVG矢量图形，包含柱状图、饼图、折线图等
- 🔒 **安全可靠**: 内置安全检查，防止XSS攻击
- 📱 **响应式设计**: 自动适配移动设备显示
- 🛠️ **简单易用**: 命令行界面，支持批量处理
- 📈 **性能优化**: 支持大文件处理和内存优化
- 🧪 **全面测试**: 完整的单元测试覆盖

## 🎯 支持格式

### 输入格式 (11种)

#### ✅ 完全支持
- **Excel 2007+**: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`
- **CSV**: `.csv` (基础格式、复杂格式、多语言)

#### ⚠️ 基础支持  
- **Excel 97-2003**: `.xls`, `.xlt`
- **Excel 二进制**: `.xlsb`  
  - 支持：数据、样式、批注、超链接、公式结构化输出  
  - 合并单元格暂不支持（输出空列表）
- **WPS Office**: `.et`, `.ett`, `.ets`  
  - 支持：数据、样式、批注、超链接、合并单元格、公式结构化输出  
  - 支持元数据、模板变量、备份信息提取

### 输出格式
- **HTML**: 完整的HTML文档或纯表格代码
- **主题**: 4种内置主题（默认、极简、暗色、打印）

## 📦 安装

### 环境要求
- Python 3.8 或更高版本
- 操作系统：Windows、macOS、Linux

### 克隆仓库
```bash
git clone https://github.com/Link-Liu/MCP-Sheet-Parser.git
cd MCP-Sheet-Parser
```

### 安装依赖
```bash
pip install pandas openpyxl xlrd pyxlsb
```

### 验证安装
```bash
python main.py --version
```

## 🚀 快速开始

### 基本使用
```bash
# 转换单个文件
python main.py data.xlsx -o output.html

# 使用不同主题
python main.py data.csv -o output.html --theme dark

# 只输出表格（不含HTML头部）
python main.py data.xlsx -o table.html --table-only

# 启用图表转换
python main.py data.xlsx -o output.html --enable-charts

# 图表转换配置
python main.py data.xlsx -o output.html --enable-charts --chart-quality high --chart-theme business

# 性能测试
python main.py data.xlsx --benchmark

# 使用CSS类而非内联样式
python main.py data.xlsx --use-css-classes
```

### 批量处理
```bash
# 批量转换多个文件
python main.py *.xlsx --batch --output-dir ./html_files/

# 批量转换并使用特定主题
python main.py data/*.csv --batch --output-dir ./output/ --theme minimal

# 使用不同性能模式
python main.py data.xlsx --performance-mode fast    # 快速模式
python main.py data.xlsx --performance-mode memory  # 内存优化模式
```

### 主题预览
```bash
# 查看所有可用主题
python main.py --list-themes
```

### 文件信息
```bash
# 查看文件信息
python main.py data.xlsx --info
```

## 🎨 主题展示

| 主题名称 | 描述 | 适用场景 |
|---------|------|----------|
| `default` | 标准样式，平衡美观与实用 | 一般用途 |
| `minimal` | 极简设计，清爽干净 | 简洁展示 |
| `dark` | 暗色主题，护眼舒适 | 夜间使用 |
| `print` | 打印优化，黑白适配 | 打印文档 |

## 📋 功能特性

### 样式与结构支持
- ✅ 字体样式（粗体、斜体、字号、颜色）
- ✅ 背景颜色
- ✅ 文本对齐（水平、垂直）
- ✅ 边框样式
- ✅ 合并单元格（WPS/Excel 2007+/97-2003，.xlsb暂不支持）
- ✅ 公式结构化输出（所有支持格式，含依赖、类型分析）

### 图表转换支持
- ✅ **图表类型**: 柱状图、条形图、折线图、饼图、面积图、散点图、组合图
- ✅ **输出格式**: SVG矢量图形，支持缩放不失真
- ✅ **图表元素**: 标题、图例、坐标轴、数据标签、网格线
- ✅ **样式配置**: 多种配色方案、字体样式、尺寸控制
- ✅ **交互功能**: 响应式设计、工具提示、动画效果
- ✅ **质量选项**: 支持低、中、高质量输出，适配不同场景

### 交互功能
- ✅ 单元格批注（悬停显示）
- ✅ 超链接（新窗口打开）
- ✅ 响应式布局
- ✅ 移动设备适配

### WPS与XLSB增强说明
- `.xlsb`：支持样式、批注、超链接、公式结构化输出，合并单元格暂不支持（输出空列表）
- `.et/.ett/.ets`：支持批注、超链接、合并单元格、样式、公式结构化输出，支持元数据、模板变量、备份信息提取

### 安全特性
- ✅ HTML内容转义
- ✅ 文件路径验证
- ✅ 大小限制检查
- ✅ 恶意内容过滤

## 📚 API使用

### 编程接口
```python
from mcp_sheet_parser.parser import SheetParser
from mcp_sheet_parser.chart_converter import ChartConverter

# 解析文件
parser = SheetParser('data.xlsb')
sheets = parser.parse()
# sheets[0]['styles'], sheets[0]['comments'], sheets[0]['hyperlinks'], sheets[0]['formula_cells']

parser = SheetParser('data.et')
sheets = parser.parse()
# sheets[0]['styles'], sheets[0]['comments'], sheets[0]['hyperlinks'], sheets[0]['merged_cells'], sheets[0]['formula_cells']

# 图表转换
chart_converter = ChartConverter()
charts = chart_converter.detect_charts_in_excel('data.xlsx')
for chart in charts:
    svg_content = chart_converter.generate_svg(chart)
    # 处理SVG内容
```

### 配置选项
```python
from mcp_sheet_parser.config import Config
from mcp_sheet_parser.config.chart import ChartConfig

config = Config()
config.INCLUDE_COMMENTS = False  # 不包含注释
config.INCLUDE_HYPERLINKS = False  # 不包含超链接
config.MAX_FILE_SIZE_MB = 50  # 最大文件大小

# 图表配置
chart_config = ChartConfig()
chart_config.ENABLE_CHART_CONVERSION = True  # 启用图表转换
chart_config.CHART_OUTPUT_FORMAT = 'svg'  # 输出SVG格式
chart_config.CHART_QUALITY = 'high'  # 高质量输出
chart_config.CHART_COLOR_SCHEME = 'business'  # 商务配色
```

## 🧪 运行测试

```bash
# 运行所有测试
python -m pytest tests/

# 运行特定测试
python -m pytest tests/test_parser.py
python -m pytest tests/test_wps_support.py
python -m pytest tests/test_xlsb_enhanced.py
```

测试覆盖范围：
- ✅ WPS格式（.et/.ett/.ets）的识别、解析、元数据、批注、超链接、合并单元格、样式、公式等
- ✅ XLSB格式的解析、样式、批注、超链接、公式等
- ✅ 所有核心功能的单元测试

## 🎯 快速演示

### 启动演示系统
```bash
# 方法1: 直接运行Python脚本
cd demo
python start_demo.py

# 方法2: Windows用户推荐使用批处理文件（解决编码问题）
cd demo
run_demo.bat

# 方法3: 测试编码是否正常
cd demo
python test_encoding.py
```

### 演示功能
演示系统提供以下功能：
- 🚀 **快速体验**: 基础功能演示
- 🎯 **完整演示**: 运行所有演示
- 📝 **创建示例**: 生成示例文件
- 📊 **图表演示**: 图表转换功能展示
- 📚 **查看文档**: 打开使用指南
- 📊 **项目总结**: 查看实现状态
- 🏠 **打开主页**: 查看主展示页面

### 编码问题解决
如果在Windows上遇到中文乱码问题：

1. **使用批处理文件**（推荐）：
   ```bash
   cd demo
   run_demo.bat
   ```

2. **手动设置编码**：
   ```bash
   chcp 65001
   set PYTHONIOENCODING=utf-8
   python demo/start_demo.py
   ```

3. **使用编码修复脚本**：
   ```bash
   cd demo
   python fix_encoding.py start_demo.py
   ```

## 📁 项目结构

```
MCP-Sheet-Parser/
├── mcp_sheet_parser/           # 核心模块
│   ├── config/                # 配置模块
│   │   ├── __init__.py        # 统一配置入口
│   │   ├── chart.py           # 图表配置
│   │   ├── file_format.py     # 文件格式配置
│   │   ├── formula.py         # 公式配置
│   │   ├── html.py            # HTML输出配置
│   │   ├── logging.py         # 日志配置
│   │   ├── performance.py     # 性能配置
│   │   └── style.py           # 样式配置
│   ├── parser.py              # 文件解析器
│   ├── html_converter.py      # HTML转换器
│   ├── formula_processor.py   # 公式处理器
│   ├── style_manager.py       # 样式管理器
│   ├── chart_converter.py     # 图表转换器
│   ├── file_processor.py      # 文件处理器
│   ├── performance.py         # 性能监控
│   ├── security.py            # 安全检查
│   ├── exceptions.py          # 异常处理
│   ├── data_validator.py      # 数据验证
│   ├── utils.py               # 工具函数
│   ├── benchmark.py           # 性能基准测试
│   └── cli.py                 # 命令行接口
├── tests/                     # 测试文件
│   ├── test_parser.py         # 解析器测试
│   ├── test_html_converter.py # HTML转换测试
│   ├── test_formula_processor.py # 公式处理测试
│   ├── test_style_manager.py  # 样式管理测试
│   ├── test_utils.py          # 工具函数测试
│   ├── test_exceptions.py     # 异常处理测试
│   ├── test_performance.py    # 性能测试
│   ├── test_config_refactor.py # 配置重构测试
│   ├── test_wps_support.py    # WPS支持测试
│   ├── test_xlsb_enhanced.py  # XLSB增强测试
│   ├── test_chart_converter.py # 图表转换测试
│   └── test_alignment_enhancement.py # 对齐增强测试
├── demo/                      # 演示系统
│   ├── start_demo.py          # 演示启动脚本
│   ├── run_demo.bat           # Windows批处理文件
│   ├── fix_encoding.py        # 编码修复工具
│   ├── test_encoding.py       # 编码测试工具
│   ├── 演示脚本/              # 演示脚本目录
│   ├── 静态展示/              # HTML展示页面
│   ├── 动态演示/              # 动态演示内容
│   ├── 文档/                  # 文档目录
│   ├── 示例文件/              # 示例文件目录
│   └── 演示总结.md            # 演示总结文档
├── 示例文件/                  # 项目示例文件
├── main.py                    # 命令行入口
└── README.md                  # 项目说明文档
```

## 📖 命令行参数

### 基本参数
- `input` - 输入文件路径（支持通配符）
- `-o` - 输出HTML文件路径
- `--output-dir, -d` - 批量处理输出目录

### 基本选项
- `input_file` - 输入文件路径（支持通配符）
- `-o, --output` - 输出HTML文件路径
- `--theme` - 选择主题（default/minimal/dark/print）
- `--table-only` - 只输出表格，不包含HTML头部
- `--encoding` - 文件编码（默认：utf-8）

### 性能选项
- `--chunk-size` - 大文件分块大小（默认：1000）
- `--max-memory` - 最大内存使用MB（默认：2048）
- `--performance-mode` - 性能模式（auto/fast/memory）
- `--disable-progress` - 禁用进度显示
- `--benchmark` - 执行性能基准测试

### 样式选项
- `--use-css-classes` - 生成CSS类而非内联样式
- `--semantic-names` - 使用语义化CSS类名
- `--template` - 样式模板（business/financial/analytics）
- `--conditional-rules` - 预定义条件格式化规则
- `--disable-conditional` - 禁用条件格式化

### 公式处理选项
- `--disable-formulas` - 禁用公式处理
- `--show-formula-text` - 在悬停时显示原始公式文本
- `--calculate-formulas` - 计算公式结果
- `--show-formula-errors` - 显示公式错误
- `--supported-functions-only` - 仅处理支持的函数

### 图表转换选项
- `--enable-charts` - 启用图表转换功能
- `--chart-format` - 图表输出格式（svg/png）
- `--chart-width` - 图表默认宽度
- `--chart-height` - 图表默认高度
- `--chart-quality` - 图表质量（low/medium/high）
- `--chart-theme` - 图表配色主题（default/business/modern/colorful）
- `--chart-responsive` - 生成响应式图表

### 信息选项
- `--verbose, -v` - 详细输出
- `--quiet, -q` - 静默模式
- `--version` - 显示版本信息

## 🔧 高级配置

### 性能优化
```python
from mcp_sheet_parser.config import UnifiedConfig

# 创建性能优化配置
config = UnifiedConfig()
config.performance.CHUNK_SIZE = 500
config.performance.MAX_MEMORY_MB = 1024
config.performance.ENABLE_PARALLEL_PROCESSING = True

# 或者使用预设配置
config = UnifiedConfig().optimize_for_performance()
```

### 质量优化
```python
# 创建质量优化配置
config = UnifiedConfig().optimize_for_quality()

# 自定义配置
config.formula.CALCULATE_FORMULAS = True
config.chart.CHART_QUALITY = 'high'
config.style.ENABLE_CONDITIONAL_FORMATTING = True
```

### 安全设置
```python
config.html.INCLUDE_COMMENTS = False      # 不包含注释
config.html.INCLUDE_HYPERLINKS = False    # 不包含超链接
config.performance.MAX_FILE_SIZE_MB = 50  # 文件大小限制
```

## 🐛 故障排除

### 常见问题

1. **编码问题**
   ```bash
   # Windows用户使用批处理文件
   cd demo
   run_demo.bat
   ```

2. **依赖缺失**
   ```bash
   pip install pandas openpyxl xlrd pyxlsb
   ```

3. **文件格式不支持**
   - 检查文件扩展名是否在支持列表中
   - 确认文件未损坏

4. **内存不足**
   - 启用流式处理：`--enable-streaming`
   - 减少最大行数/列数限制

### 错误代码
- `0`: 成功
- `1`: 一般错误
- `2`: 文件验证失败
- `3`: 严重错误
- `130`: 用户中断

## 📈 性能基准

| 文件类型 | 大小 | 处理时间 | 内存使用 | 支持功能 |
|---------|------|----------|----------|----------|
| Excel 2007+ | 1MB | ~2秒 | ~50MB | 完整功能+图表 |
| CSV | 1MB | ~1秒 | ~30MB | 基础功能 |
| WPS | 1MB | ~3秒 | ~60MB | 增强功能 |
| XLSB | 1MB | ~2.5秒 | ~55MB | 基础功能 |
| Excel 97-2003 | 1MB | ~2.5秒 | ~45MB | 基础功能 |

### 图表转换性能

| 图表类型 | 复杂度 | 处理时间 | 输出大小 | 质量 |
|---------|--------|----------|----------|------|
| 柱状图 | 简单 | ~0.5秒 | ~10KB | 高 |
| 饼图 | 简单 | ~0.3秒 | ~8KB | 高 |
| 折线图 | 中等 | ~0.8秒 | ~15KB | 高 |
| 复杂组合图 | 复杂 | ~1.5秒 | ~25KB | 高 |


## 🙏 致谢

- [pandas](https://pandas.pydata.org/) - 数据处理
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel 2007+ 支持
- [xlrd](https://xlrd.readthedocs.io/) - Excel 97-2003 支持
- [pyxlsb](https://github.com/willtrnr/pyxlsb) - XLSB 支持

## 📞 联系方式
2256981152@qq.com

**MCP-Sheet-Parser** - 让表格转换变得简单高效！ 🚀
